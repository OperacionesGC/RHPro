Attribute VB_Name = "MdlTiposBusquedas"
    Option Explicit

' ---------------------------------------------------------------------------------------------
' Descripcion:Modulo de Busquedas
' Autor      :FGZ
' Fecha      :05/08/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

' Procedimienos de Busqueda de Par�metros
Public Sub Guardar_Nov_Hist()
   
    guarda_nov = True

End Sub

Public Sub Buscar(p_valor As Single, p_retro As Date, p_bien As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Veificar si est� en el cache.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
If objCache.EsSimboloDefinido(CStr(NroConce)) Then
    p_valor = objCache.Valor(CStr(NroConce))
    p_retro = CDate("")
    p_bien = True
Else
    ' Ejecutar la busqueda
    Call bus_Acum
    
    ' Insertar la informacion en el cahce
    If Bien Then
        objCache.Insertar_Simbolo CStr(NroConce), Valor
    End If
End If


End Sub

Public Sub EjecutarBusqueda(ByVal tipoBus As Long, ByVal concnro As Long, ByVal prog As Long, ByRef val As Single, ByRef Fecha As Date, ByRef Ok As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Simula  el RUN VALUE (programa.progarch) (buliq-cabliq.empleado,nrogrupo,buliq-concepto.concnro,for_tpa.tpanro,buliq-cabliq.cliqnro, OUTPUT val,OUTPUT fec, OUTPUT ok) NO-ERROR.
'              Este es el procedimiento llamador dependiendo del tipo de Busqueda
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim nombre As String
Dim Rastreo As Boolean

Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer

Dim TpoInicialBus
Dim TpoFinalBus

    
    TpoInicialBus = GetTickCount

NroProg = prog

Select Case tipoBus
Case 1: ' Det. Autom�tica - Novedad
            ' ???????????????????????
    nombre = "Det. Autom�tica - Novedad"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & " no implementada "
    End If
    
Case 2:     'B�squeda interna
    nombre = "B�squeda interna"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_interna
Case 3:     'Escalas
    nombre = "Escalas"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Grilla(tipoBus, concnro, prog)
Case 4:     'Acumulador Liq. Actual
    nombre = "Acumulador Liq. Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Acum0
Case 5:     'Concepto Liq. Actual
    nombre = "Concepto Liq. Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Concep0
Case 6:     'Fc. S/ acumulador procesos
    nombre = "Fc. S/ acumulador procesos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Acum
Case 7:     'Acum. Mensual meses fijos
    nombre = "Acum. Mensual meses fijos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Acum3
Case 8:     'Acum. Mensual meses variables
    nombre = "Acum. Mensual meses variables"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Acum4
Case 9:     'Novedad de otro concepto
    nombre = "Novedad de otro concepto"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Concep1
Case 10:    'Antiguedad del empleado
    nombre = "Antiguedad del empleado"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Anti0(antdia, antmes, antanio)
Case 11:    'Remuneraci�n p/Ganancias
    nombre = "Remuneraci�n p/Ganancias"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Remun0
Case 12:    'Prom. Asig. Familiares
    nombre = "Prom. Asig. Familiares"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Acum5
Case 13:    'Campo de la Base de Datos
    nombre = "Campo de la Base de Datos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Campo0
Case 14:    'Prestamos
    nombre = "Prestamos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Prestamos
Case 15:    'PreCalculo de parametros
    nombre = "PreCalculo de parametros"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Tcotpa0
Case 16:    'Conceptos Mese Fijos
    nombre = "Conceptos Mese Fijos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Concep3
Case 17:    'no existe
    nombre = "no existe"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
Case 18:    'Calculo de Vacaciones
    nombre = "Calculo de Vacaciones"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Anti0(antdia, antmes, antanio)
Case 19:    'Calculo de indemnizaci�n
    nombre = "Calculo de indemnizaci�n"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Anti0(antdia, antmes, antanio)
Case 20:    'Cotizaci�n de Monedas
    nombre = "Cotizaci�n de Monedas"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Cotmon1
Case 21:    'Acumuladores Imponibles Mensuales
    nombre = "Acumuladores Imponibles Mensuales"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_ImponiblesMensuales
Case 22:    'Embargos
    nombre = "Embargos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Embargos
Case 23:    'Descuentos de Mutuales y Sindicatos
    nombre = "Descuentos de Mutuales y Sindicatos"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    'Call bus_
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & " no implementada "
    End If
    
Case 24:    'Estructuras a una Fecha
    nombre = "Estructuras a una Fecha"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Estructura
Case 25:    'Antiguedad en la estructura
    nombre = "Antiguedad en la estructura"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_AntEstructura
Case 26:    'Zona de Domicilio
    nombre = "Zona de Domicilio"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_ZonaDom
Case 27:    'Pago / Descuento de Licencias
    nombre = "Pago / Descuento de Licencias"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_PagoDtoLic
Case 28:    'Novedades de GTI
    nombre = "Novedades de GTI"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_NovGTI
Case 29:    ' Imponibles del proceso
    nombre = "Acumuladores Imponibles del Proceso Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_ImponiblesDelProceso
Case 30: 'Dias habiles del Mes
    nombre = "Calculo de dias habiles entre dos fechas"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasHabiles
Case 31: 'Licencias
    nombre = "Licencias"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Licencias
Case 32: 'Vales
    nombre = "Vales"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Vales
Case 33: 'Valores Constantes
    nombre = "Valores Constantes"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Constantes
Case 34: 'Parametro de un Concepto
    nombre = "Parametro de un concepto"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_ParametroConcepto
Case 35: 'Remuneracion del Empleado
    nombre = "Remuneracion del Empleado"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_SueldoRemun
Case 36: 'Mes Liquidacion Actual
    nombre = "Mes Liquidacion Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Sis_MesActual
Case 37: 'A�o Liquidacion Actual
    nombre = "A�o Liquidacion Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Sis_AnioActual
Case 38: 'Semestre Liquidacion Actual
    nombre = "Semestre Liquidacion Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Sis_SemestreActual
Case 39: 'Modelo de la Liquidacion Actual
    nombre = "Modelo de la Liquidacion Actual"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Sis_ModeloLiqActual
Case 40: 'Dias del mes
    nombre = "Cantidad de dias del mes"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Sis_Dias_Mes
Case 41: 'Dias Correspondientes de Vacaciones por escala
    nombre = "Dias Correspondientes de Vac"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasVac
Case 42: 'Dias Correspondientes de Vacaciones
    nombre = "Dias Correspondientes de Vac por vacdiascor"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasVac2
Case 43: 'Licencias sin Tope, es decir, por mes calendario
    nombre = "Licencias Mes Calendario"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_LicenciasMesCalendario
Case 44: 'Dias Mes Calendario en la estructura
    nombre = "Dias Mes Calendario en la estructura"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasMesCalendario_enEstructura
Case 45: 'Asignaciones Familiares
    nombre = "Asignaciones Familiares"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_AsignacionesFliares
Case 46: 'Dias habiles Mes otra liquidacion
    nombre = "Calculo de dias habiles entre dos fechas correspondientes a meses anteriores o posteriores a la liq actual "
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasHabilesMesLiquidacion
Case 47: 'Promedio de Vacaciones
    nombre = "Promedio de Vacaciones"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_PromedioVacaciones
Case 48: 'Licencias segun periodo GTI
    nombre = "Licencias segun periodo GTI"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_LicenciasPeriodoGTI
Case 49: 'Dias para SAC
    nombre = "Dias para SAC"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasSAC
Case 50: 'Vacaciones no Gozadas
    nombre = "Vacaciones no Gozadas"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Vac_No_Gozadas(concnro, prog)
Case 51: 'Vacaciones no Gozadas Pendientes
    nombre = "Vacaciones no Gozadas Pendientes"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Vac_No_Gozadas_Pendientes(concnro, prog)
Case 52: 'Edad del Empleado
    nombre = "Edad del Empleado"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call Bus_Edad_Empleado
Case 53: 'SAC Proporcional
    nombre = "SAC Proporcional"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasSAC_Proporcional
Case 54: 'Dias de Ingreso
    nombre = "Dias de Ingreso"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasDeIngreso
Case 55: 'Base Licencias
    nombre = "Base Licencias"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_BaseLicencias
Case 56: 'Acumulador o concepto en otro Legajo
    nombre = "Acumulador o concepto en otro Legajo"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_ValorEnOtroLegajo
Case 57: 'Dias en Mes Segun Fase
    nombre = "Dias en Mes Segun Fase"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasEnMesSegunFase
Case 58: 'Antiguedad segun acumulador
    nombre = "Antiguedad segun acumulador"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Antiguedad_Por_Acumulador
Case 59: 'Vacaciones no Gozadas a Pagar
    nombre = "Vacaciones no Gozadas a Pagar"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Vac_No_Gozadas_A_Pagar(concnro, prog)
Case 60: 'Horas a Justificar
    nombre = "Horas a Justificar"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Licencias_Horas_A_Justificar
Case 61: 'Feriados en Quincena
    nombre = "Feriados en Quincena"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Feriados_Quincena
Case 62: 'Dias Correspondientes de Vacaciones(Antiguedad a Fecha de alta)
    nombre = "Dias Correspondientes de Vacaciones(Antiguedad a Fecha de alta)"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasVac_Antig2
Case 63: 'Mes de Fecha de pago del proceso de liq
    nombre = "Mes de Fecha de pago del proceso de liq"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_MesPagoProceso
Case 64: 'Dias habiles con Fases
    nombre = "Calculo de dias habiles entre dos fechas con fases"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_DiasHabiles_ConFases
Case 65: 'Partes Diario
    nombre = "Calculo de partes para un periodo"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Partes_Diarios
Case 66: 'BAE
    nombre = "Calculo de BAE para  un periodo"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_BAE
Case 67: 'Movilidad
    nombre = "Calculo de movilidad para un proceso"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Movilidad
    
Case 68: 'Busqueda de escala, tiene que hacer nada
    nombre = "Busqueda de escala"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If

Case 69: 'Cantidad de empleados en la estructura
    nombre = "Calculo de empleados en la estructura"
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Tipo de busqueda " & nombre
    End If
    Call bus_Cant_Empl_Estr
    
Case Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & " Tipo de busqueda desconocido " & tipoBus
    End If
End Select

val = Valor
Ok = Bien

'' FGZ - para sacar estadisticas de las busquedas ejecutadas
'' Este codigo no va activo. es solo para testear
'Rastreo = True
'If Rastreo Then
'    If Borrar_Estadisticas Then
'        StrSql = "DELETE FROM tipobus_Ejecutadas"
'        objConn.Execute StrSql, , adExecuteNoRecords
'
'        Borrar_Estadisticas = False
'    End If
'
'    TpoFinalBus = GetTickCount
'    ' Inserto en TipoBusEjecutadas
'    StrSql = "INSERT INTO tipobus_Ejecutadas (" & _
'             "tipobusnro,nombre,concnro,prognro,exito,fechaej,tiempo_ej" & _
'             ") VALUES (" & tipoBus & _
'             ",'" & Left(nombre, 30) & _
'             "'," & Concnro & _
'             "," & prog & _
'             "," & CInt(ok) & _
'             "," & ConvFecha(Date) & _
'             "," & (TpoFinalBus - TpoInicialBus) & _
'             " )"
'    objConn.Execute StrSql, , adExecuteNoRecords
'End If

' SQL para ver los resultados
'SELECT     tipobusnro, COUNT(tipobusnro) AS Cantidad, SUM(tiempo_ej) AS Tiempo
'From tipobus_ejecutadas
'Where (tiempo_ej <> 0) And (exito = -1)
'GROUP BY tipobusnro
'ORDER BY tipobusnro

End Sub


' --------------------------------------------------------
' Procedimientos Espec�ficos de las Busquedas
' --------------------------------------------------------



Public Sub bus_Concep0()
' ---------------------------------------------------------------------------------------------
' Descripcion: Concepto de liquidacion actual. gconcep0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroConc As Long         'concepto.concnro
Dim Oblig As Boolean        ' Obligatorio retornar valor
Dim cant As Integer         ' 1 - Cantidad
                            ' 2 - Monto
                            
'Dim Param_cur As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset

    Bien = False
    Valor = 0

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroConc = Arr_Programa(NroProg).Auxint2
        Oblig = CBool(Arr_Programa(NroProg).Auxlog1)
        cant = Arr_Programa(NroProg).Auxint1
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

    ' FGZ - 09/02/2004
    'Busco en el cache de concepto del empleado
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco en el cache de concepto del empleado "
    End If
    
    If objCache_detliq_Monto.EsSimboloDefinido(CStr(NroConc)) Then
        If cant = 1 Then
            Valor = objCache_detliq_Cantidad.Valor(CStr(NroConc))
        Else
            Valor = objCache_detliq_Monto.Valor(CStr(NroConc))
        End If
        Bien = True
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "encontr� en Cache con valor " & Valor
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No encontr� en Cache "
        End If
        
        If CBool(Oblig) Then
            Bien = False
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Concepto " & NroConc & " no encontrado "
                Flog.writeline Espacios(Tabulador * 4) & "Retorna Falso porque es obligatorio "
            End If
        Else
            Valor = 0
            Bien = True
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Concepto " & NroConc & " no encontrado "
                Flog.writeline Espacios(Tabulador * 4) & "Retorna Verdadero y cero porque no es obligatorio "
            End If
        End If
    End If

'    StrSql = "SELECT dlicant, dlimonto FROM detliq " & _
'                 " WHERE cliqnro =" & buliq_cabliq!cliqnro & _
'                 " AND concnro =" & NroConc
'        OpenRecordset StrSql, rs_Detliq
'
'    If Not rs_Detliq.EOF Then
'        If cant = 1 Then
'            Valor = rs_Detliq!dlicant
'        Else
'            Valor = rs_Detliq!dlimonto
'        End If
'        Bien = True
'    Else
'        Bien = False
'    End If
            

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
'Set Param_cur = Nothing
Set rs_Detliq = Nothing

End Sub


Public Sub bus_Concep0_old()
' ---------------------------------------------------------------------------------------------
' Descripcion: Concepto de liquidacion actual. gconcep0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroConc As Long         'concepto.concnro
Dim Oblig As Boolean        ' Obligatorio retornar valor
Dim cant As Integer         ' 1 - Cantidad
                            ' 2 - Monto
                            
'Dim Param_cur As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset

    Bien = False
    Valor = 0

    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroConc = Arr_Programa(NroProg).Auxint2
        Oblig = CBool(Arr_Programa(NroProg).Auxlog1)
        cant = Arr_Programa(NroProg).Auxint1
    Else
        Exit Sub
    End If

    StrSql = "SELECT dlicant, dlimonto FROM detliq " & _
                 " WHERE cliqnro =" & buliq_cabliq!cliqnro & _
                 " AND concnro =" & NroConc
        OpenRecordset StrSql, rs_Detliq
        
    If Not rs_Detliq.EOF Then
        If cant = 1 Then
            Valor = rs_Detliq!dlicant
        Else
            Valor = rs_Detliq!dlimonto
        End If
        Bien = True
    Else
        Bien = False
    End If
            

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
'Set Param_cur = Nothing
Set rs_Detliq = Nothing

End Sub

Public Sub bus_Concep3()
' ---------------------------------------------------------------------------------------------
' Descripcion: conceptos meses fijos. gconcep3.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroConc As Long         'concepto.concnro
Dim Oblig As Boolean        ' Obligatorio
Dim cant As Boolean         ' true - Cantidad
                            ' false - Monto
Dim Semestral As Integer    ' 1 - Semestral
                            ' 2 - Anual
                            ' 3 - Meses
Dim Meses As Integer
Dim Opcion As Long          ' 1 - Sumatoria
                            ' 2 - Maximo
                            ' 3 - Promedio
                            ' 4 - Promedio sin 0
                            ' 5 - Minimo

Dim Incluye As Integer      ' 0  - No Incluye
                            ' 1  - Proceso Actual
                            ' 2  - Periodo Actual con Proceso actual
                            ' 3  - Periodo Actual sin proceso actual
                            
Dim MesDeInicioSemestre As Integer 'Mes de Inicio (en caso de que sea Semestral)

Dim nro_desde As Long
Dim nro_hasta As Long
Dim Cant_per As Single
Dim CantMeses As Integer
Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim i As Integer
Dim pliqAnterior As Long

Dim UsaActual As Boolean
Dim PrimeraVes As Boolean
Dim aux As String

'Dim Param_cur As New ADODB.Recordset
Dim rs_Periodos As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset

Bien = False
Valor = 0

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Meses = CInt(Arr_Programa(NroProg).Auxint1)
        NroConc = CLng(Arr_Programa(NroProg).Auxint2)
        Opcion = CInt(Arr_Programa(NroProg).Auxint3)
        Semestral = CInt(Arr_Programa(NroProg).Auxint4)
        Incluye = Arr_Programa(NroProg).Auxint5
        Oblig = CBool(Arr_Programa(NroProg).Auxlog1)
        cant = CBool(Arr_Programa(NroProg).Auxlog2)
        If Semestral = 1 Then
            MesDeInicioSemestre = Arr_Programa(NroProg).Auxchar1
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros cargados "
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If

    If CBool(USA_DEBUG) Then
        Select Case Semestral
        Case 1:
            aux = "semestral, Mes de inicio: " & MesDeInicioSemestre
        Case 2:
            aux = "Anual "
        Case 3:
            aux = " mensual (" & Meses & " meses) "
        End Select
        Flog.writeline Espacios(Tabulador * 4) & "Busqueda " & aux
        aux = ""
        Flog.writeline Espacios(Tabulador * 4) & "Concepto: " & NroConc
        
        Select Case Opcion
        Case 1:
            aux = "Sumatoria"
        Case 2:
            aux = "Maximo"
        Case 3:
            aux = "Promedio"
        Case 4:
            aux = "Promedio sin 0"
        Case 5:
            aux = "Minimo"
        End Select
        Flog.writeline Espacios(Tabulador * 4) & "Operacion: " & aux
        aux = ""
        
        Select Case Incluye
        Case 0:
            aux = "No Incluye el periodo actual ni proceso actual"
        Case 1:
            aux = "No Incluye el periodo actual pero si el proceso actual"
        Case 2:
            aux = "Incluye el periodo actual y proceso actual"
        Case 3:
            aux = "Incluye el periodo actual pero no el proceso actual"
        End Select
        Flog.writeline Espacios(Tabulador * 4) & aux
        aux = ""
        
        If Oblig Then
            aux = "Es Obligatorio retornar "
            If cant Then
                aux = aux & "cantidad "
            Else
                aux = aux & "Monto "
            End If
        Else
            aux = "No es Obligatorio retornar "
            If cant Then
                aux = aux & "cantidad "
            Else
                aux = aux & "Monto "
            End If
        End If
        Flog.writeline Espacios(Tabulador * 4) & aux
    End If

Select Case Semestral
Case 1: ' Semestral
    MesDesde = MesDeInicioSemestre
    AnioDesde = buliq_periodo!pliqanio
    Select Case Incluye
    Case 0: 'Semestre actual y No Incluye ni Proceso actual ni periodo actual
        If buliq_periodo!pliqmes > 6 Then
            If MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 7 - 1) + (MesDeInicioSemestre - 7)
            Else
                Meses = (buliq_periodo!pliqmes - 7 - 1)
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            If Not MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 1 - 1) + (MesDeInicioSemestre - 1)
            Else
                Meses = (buliq_periodo!pliqmes - 1 - 1)
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
    Case 1: 'Semestre actual y Incluye el proceso actual pero no el periodo actual
        If buliq_periodo!pliqmes > 6 Then
            If MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 7 - 1) + (MesDeInicioSemestre - 7)
            Else
                Meses = (buliq_periodo!pliqmes - 7 - 1)
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            If Not MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 1 - 1) + (MesDeInicioSemestre - 1)
            Else
                Meses = (buliq_periodo!pliqmes - 1 - 1)
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = True
    Case 2: 'Semestre actual y Incluye el Periodo actual y el proceso actual
        If buliq_periodo!pliqmes > 6 Then
            If MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 7) + (MesDeInicioSemestre - 7)
            Else
                Meses = (buliq_periodo!pliqmes - 7)
            End If
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        Else
            If Not MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 1) + (MesDeInicioSemestre - 1)
            Else
                Meses = (buliq_periodo!pliqmes - 1)
            End If
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = True
    Case 3: 'Semestre actual y Incluye el Periodo actual pero no el proceso actual
        If buliq_periodo!pliqmes > 6 Then
            If MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 7) + (MesDeInicioSemestre - 7)
            Else
                Meses = (buliq_periodo!pliqmes - 7)
            End If
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        Else
            If Not MesDeInicioSemestre > 6 Then
                Meses = (buliq_periodo!pliqmes - 1) + (MesDeInicioSemestre - 1)
            Else
                Meses = (buliq_periodo!pliqmes - 1)
            End If
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
        
    Case 4: 'Semestre anterior
        If buliq_periodo!pliqmes > 6 Then
            If MesDeInicioSemestre > 6 Then
                Meses = 5 + (MesDeInicioSemestre - 1)
            Else
                Meses = 5
            End If
            MesHasta = 6
            AnioHasta = buliq_periodo!pliqanio
        Else
            If Not MesDeInicioSemestre > 6 Then
                Meses = 7 + (MesDeInicioSemestre - 7)
            Else
                Meses = 7
            End If
            MesHasta = 12
            AnioHasta = buliq_periodo!pliqanio - 1
        End If
        UsaActual = False
    End Select
Case 2: ' Anual
    MesDesde = 1
    AnioDesde = buliq_periodo!pliqanio
    Select Case Incluye
    Case 0: 'A�o actual y No Incluye ni Proceso actual ni periodo actual
            Meses = buliq_periodo!pliqmes - 1 - 1
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
            UsaActual = False
    Case 1: 'A�o actual y Incluye el proceso actual pero no el periodo actual
            Meses = buliq_periodo!pliqmes - 1 - 1
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
            UsaActual = True
    Case 2: 'A�o actual y Incluye el Periodo actual y el proceso actual
            Meses = buliq_periodo!pliqmes - 1
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
            UsaActual = True
    Case 3: 'A�o actual y Incluye el Periodo actual pero no el proceso actual
            Meses = buliq_periodo!pliqmes - 1
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
            UsaActual = False
    Case 4: 'A�o anterior
            Meses = 11
            MesHasta = 12
            AnioHasta = buliq_periodo!pliqanio - 1
            UsaActual = False
    End Select

Case 3:     ' Meses
'    nro_desde = buliq_periodo!pliqnro - 1
'    nro_hasta = buliq_periodo!pliqnro - 1   ' para tener un default
    
    Select Case Incluye
    Case 0: 'No Incluye ni Proceso actual ni periodo actual
        If buliq_periodo!pliqmes >= Meses Then
            MesDesde = buliq_periodo!pliqmes - Meses
            AnioDesde = buliq_periodo!pliqanio
            If MesDesde = 0 Then
                MesDesde = 12
                AnioDesde = buliq_periodo!pliqanio - 1
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            MesDesde = 12 - (buliq_periodo!pliqmes - Meses)
            MesHasta = buliq_periodo!pliqmes - 1
            AnioDesde = buliq_periodo!pliqanio - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
    Case 1: 'Incluye el proceso actual pero no el periodo actual
        If buliq_periodo!pliqmes >= Meses Then
            MesDesde = buliq_periodo!pliqmes - Meses
            AnioDesde = buliq_periodo!pliqanio
            If MesDesde = 0 Then
                MesDesde = 12
                AnioDesde = buliq_periodo!pliqanio - 1
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            MesDesde = 12 - (buliq_periodo!pliqmes - Meses)
            MesHasta = buliq_periodo!pliqmes - 1
            AnioDesde = buliq_periodo!pliqanio - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = True
    Case 2: 'Incluye el Periodo actual y el proceso actual
        If buliq_periodo!pliqmes >= Meses Then
            MesDesde = buliq_periodo!pliqmes - Meses
            AnioDesde = buliq_periodo!pliqanio
            If MesDesde = 0 Then
                MesDesde = 12
                AnioDesde = buliq_periodo!pliqanio - 1
            End If
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        Else
            MesDesde = 12 - (Meses - buliq_periodo!pliqmes + 1)
            MesHasta = buliq_periodo!pliqmes
            AnioDesde = buliq_periodo!pliqanio - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
    Case 3: 'Incluye el Periodo actual y no el proceso actual
        If buliq_periodo!pliqmes >= Meses Then
            MesDesde = buliq_periodo!pliqmes - Meses
            AnioDesde = buliq_periodo!pliqanio
            If MesDesde = 0 Then
                MesDesde = 12
                AnioDesde = buliq_periodo!pliqanio - 1
            End If
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        Else
            MesDesde = 12 - (Meses - buliq_periodo!pliqmes + 1)
            MesHasta = buliq_periodo!pliqmes
            AnioDesde = buliq_periodo!pliqanio - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
    Case 4: 'no disponible (no tiene sentido)
        If buliq_periodo!pliqmes >= Meses Then
            MesDesde = buliq_periodo!pliqmes - Meses
            AnioDesde = buliq_periodo!pliqanio
            If MesDesde = 0 Then
                MesDesde = 12
                AnioDesde = buliq_periodo!pliqanio - 1
            End If
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            MesDesde = 12 - (buliq_periodo!pliqmes - Meses)
            MesHasta = buliq_periodo!pliqmes - 1
            AnioDesde = buliq_periodo!pliqanio - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
    End Select
End Select


If CBool(USA_DEBUG) Then
    Flog.writeline
    Flog.writeline Espacios(Tabulador * 4) & "Busco todos los detliq "
    Flog.writeline Espacios(Tabulador * 4) & "Desde el mes  " & MesDesde & " del a�o " & AnioDesde
    Flog.writeline Espacios(Tabulador * 4) & "hasta el mes  " & MesHasta & " del a�o " & AnioHasta
End If

' Busco todos lod detliq entre los meses
StrSql = "SELECT * FROM periodo " & _
         " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
         " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
         " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
         " WHERE cabliq.empleado = " & NroEmple & _
         " AND ((" & MesDesde & " <= periodo.pliqmes AND periodo.pliqanio >= " & AnioDesde & ")" & _
         " AND (periodo.pliqanio <= " & AnioHasta & " AND periodo.pliqmes <= " & MesHasta & "))" & _
         " AND detliq.concnro = " & NroConc & _
         " AND proceso.pronro <> " & buliq_proceso!pronro & _
         " ORDER by periodo.pliqnro"
OpenRecordset StrSql, rs_Periodos

Cant_per = 0
pliqAnterior = 0

If Opcion = 5 Then ' Minimo
    If Not rs_Periodos.EOF Then
        rs_Periodos.MoveFirst
        
        If cant Then
            Valor = rs_Periodos!dlicant
        Else
            Valor = rs_Periodos!dlimonto
        End If
        rs_Periodos.MoveNext
    End If
End If

Do While Not rs_Periodos.EOF
    Select Case Opcion
    Case 1:     'Sumatoria
        If cant Then
            Valor = Valor + rs_Periodos!dlicant
        Else
            Valor = Valor + rs_Periodos!dlimonto
        End If
    
    Case 2:     ' Maximo
        If cant Then
            If rs_Periodos!dlicant > Valor Then
                Valor = rs_Periodos!dlicant
            End If
        Else
            If rs_Periodos!dlimonto > Valor Then
                Valor = rs_Periodos!dlimonto
            End If
        End If
    
    Case 3:     ' Promedio
        ' Cuento la cantidad de periodos
        ' if first-of(periodo.pliqnro) tehn
        If pliqAnterior <> rs_Periodos!PliqNro Then
            pliqAnterior = rs_Periodos!PliqNro
            Cant_per = Cant_per + 1
        End If
        If cant Then
            Valor = Valor + rs_Periodos!dlicant
        Else
            Valor = Valor + rs_Periodos!dlimonto
        End If
        If Cant_per = 0 Then
            Valor = 0
        Else
            Valor = Valor / Cant_per
        End If
    Case 4:     ' Promedio si 0
        ' Cuento la cantidad de periodos
        ' if first-of(periodo.pliqnro) tehn
        If pliqAnterior <> rs_Periodos!PliqNro Then
            pliqAnterior = rs_Periodos!PliqNro
            
            If cant Then
                If rs_Periodos!dlicant <> 0 Then
                    Cant_per = Cant_per + 1
                End If
            Else
                If rs_Periodos!dlimonto <> 0 Then
                    Cant_per = Cant_per + 1
                End If
            End If
        End If
        If cant Then
            Valor = Valor + rs_Periodos!dlicant
        Else
            Valor = Valor + rs_Periodos!dlimonto
        End If
        
        If Cant_per = 0 Then
            Valor = 0
        Else
            Valor = Valor / Cant_per
        End If
    
    Case 5:     ' Minimo
        If cant Then
            If rs_Periodos!dlicant < Valor Then
                Valor = rs_Periodos!dlicant
            End If
        Else
            If rs_Periodos!dlimonto < Valor Then
                Valor = rs_Periodos!dlimonto
            End If
        End If
    
    Case Else
    End Select
    
    ' Siguiente registro
    rs_Periodos.MoveNext
Loop


' Si tiene en cuenta la liquidacion actual
If UsaActual Then
    StrSql = "SELECT * FROM detliq " & _
             " INNER JOIN cabliq ON cabliq.cliqnro = detliq.cliqnro " & _
             " WHERE cabliq.empleado = " & NroEmple & _
             " AND cabliq.cliqnro = " & buliq_cabliq!cliqnro & _
             " AND detliq.concnro = " & NroConc & _
             " ORDER by cabliq.cliqnro"
    OpenRecordset StrSql, rs_Detliq


PrimeraVes = True
Do While Not rs_Detliq.EOF
    Select Case Opcion
    Case 1:     'Sumatoria
        If cant Then
            Valor = Valor + rs_Detliq!dlicant
        Else
            Valor = Valor + rs_Detliq!dlimonto
        End If
    
    Case 2:     ' Maximo
        If cant Then
            If rs_Detliq!dlicant > Valor Then
                Valor = rs_Detliq!dlicant
            End If
        Else
            If rs_Detliq!dlimonto > Valor Then
                Valor = rs_Detliq!dlimonto
            End If
        End If
    
    Case 3: ' Promedio
        If PrimeraVes Then
            Cant_per = Cant_per + 1
            PrimeraVes = False
        End If
        If cant Then
            Valor = Valor + rs_Detliq!dlicant
        Else
            Valor = Valor + rs_Detliq!dlimonto
        End If
        If Cant_per = 0 Then
            Valor = 0
        Else
            Valor = Valor / Cant_per
        End If
    Case 4: ' Promedio si 0
        If PrimeraVes Then
            PrimeraVes = False
            
            If cant Then
                If rs_Detliq!dlicant <> 0 Then
                    Cant_per = Cant_per + 1
                End If
            Else
                If rs_Detliq!dlimonto <> 0 Then
                    Cant_per = Cant_per + 1
                End If
            End If
        End If
        If cant Then
            Valor = Valor + rs_Detliq!dlicant
        Else
            Valor = Valor + rs_Detliq!dlimonto
        End If
        
        If Cant_per = 0 Then
            Valor = 0
        Else
            Valor = Valor / Cant_per
        End If
    
    Case 5:     ' Minimo
        If cant Then
            If rs_Detliq!dlicant < Valor Then
                Valor = rs_Detliq!dlicant
            End If
        Else
            If rs_Detliq!dlimonto < Valor Then
                Valor = rs_Detliq!dlimonto
            End If
        End If
    
    Case Else
    End Select
    
    ' Siguiente registro
    rs_Detliq.MoveNext
Loop

End If

Bien = True

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_Periodos.State = adStateOpen Then rs_Periodos.Close
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
'Set Param_cur = Nothing
Set rs_Periodos = Nothing
Set rs_Detliq = Nothing

End Sub


Public Sub bus_Acum()
' Fc. S/ acumulador Procesos

Dim NroAcu As Long      ' Nro de Acumulador
Dim Meses As Integer    ' Ultimos X Mese
Dim cant As Boolean     ' True  - Cantidad
                        ' False - Monto
Dim Opcion As Long      ' 1 - Sumatoria
                        ' 2 - Maximo
                        ' 3 - Promedio
Dim tipo_mes As Boolean ' True  = Mes Anterior a la Licencia
                        ' False = Mes Anterior a la Liquidacion Actual

'Dim Param_cur As New ADODB.Recordset
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        Meses = Arr_Programa(NroProg).Auxint2
        cant = CBool(Arr_Programa(NroProg).Auxlog2)
        Opcion = Arr_Programa(NroProg).Auxint3
        tipo_mes = CBool(Arr_Programa(NroProg).Auxlog1)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "SEGUN CONDICION LLAMA A bus_acum1 o bus_acum2 "
End If

'SEGUN CONDICION LLAMA A bus_acum1 o bus_acum2
If tipo_mes Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Mes Anterior a la Licencia. Llama a bus_acum1 "
    End If
    Call bus_Acum1(NroAcu, Meses, cant)
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Mes Anterior a la Liquidacion Actual. Llama a bus_acum1 "
    End If
    Call bus_Acum2(NroAcu, Meses, cant, Opcion, tipo_mes)
End If

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

End Sub

Public Sub bus_Acum0()
' ---------------------------------------------------------------------------------------------
' Descripcion: Concepto de liquidacion actual. gacum0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroAcu As Long          ' acumulador.acunro
Dim Oblig As Boolean        ' Obligatorio retornar valor
Dim cant As Integer         ' 1 - Cantidad
                            ' 2 - Monto
'Dim Param_cur As New ADODB.Recordset
'Dim rs_AcuLiq As New ADODB.Recordset

    
    Bien = False
    Valor = 0

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If

    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        Oblig = CBool(Arr_Programa(NroProg).Auxlog1)
        cant = Arr_Programa(NroProg).Auxint2
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada "
        End If
        Exit Sub
    End If

    'FGZ  - 09/02/204
    'desde ahora busco en el cache(si no est� es que no se liquid�)
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco en el cache(si no est� es que no se liquid�) "
    End If

    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(NroAcu)) Then
        If cant = 1 Then
            Valor = objCache_Acu_Liq_Cantidad.Valor(CStr(NroAcu))
        Else
            Valor = objCache_Acu_Liq_Monto.Valor(CStr(NroAcu))
        End If
        Bien = True
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Encontr� en el cache con valor " & Valor
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No encontr� en el cache"
        End If
    
        If Oblig Then
            Bien = False
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Acumulador " & NroAcu & " no encontrado y es obligatorio. Retorna Falso. "
            End If
        Else
            Bien = True
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Acumulador " & NroAcu & " no encontrado pero no es obligatorio. Retorna Verdadero. "
            End If
        End If
    End If

'    StrSql = "SELECT alcant, almonto FROM acu_liq " & _
'                 " WHERE cliqnro =" & buliq_cabliq!cliqnro & _
'                 " AND acunro =" & NroAcu
'        OpenRecordset StrSql, rs_AcuLiq
'
'    If Not rs_AcuLiq.EOF Then
'        If cant = 1 Then
'            Valor = rs_AcuLiq!alcant
'        Else
'            Valor = rs_AcuLiq!almonto
'        End If
'        Bien = True
'    Else
'        Bien = False
'    End If
            

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'If rs_AcuLiq.State = adStateOpen Then rs_AcuLiq.Close
'Set Param_cur = Nothing
'Set rs_AcuLiq = Nothing

End Sub



Public Sub bus_Acum1(NroAcu As Long, Meses As Integer, cant As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca los acumuladores de la licencia con cantidad o monto. gacum1.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
                     
Dim nro_desde As Long
Dim nro_hasta As Long
Dim i As Integer
Dim Cant_per As Single
Dim fecinicio As Date

Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim RS_Acu As New ADODB.Recordset
    
    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Controlo que la licencia este entre las fechas del periodo o "
        Flog.writeline Espacios(Tabulador * 4) & "que la licencia termine en el periodo o "
        Flog.writeline Espacios(Tabulador * 4) & "que la licencia empiece en el periodo o "
        Flog.writeline Espacios(Tabulador * 4) & "que la licencia empiece antes y no termine en el periodo"
    End If
  
    ' Controlar que la licencia este entre las fechas del periodo o
    ' que la licencia termine en el periodo o
    ' que la licencia empiece en el periodo o
    ' que la licencia empiece antes y no termine en el periodo
    StrSql = "SELECT * FROM emp_lic INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
             " WHERE empleado = " & NroEmple & _
             " AND elfechadesde <= " & ConvFecha(buliq_periodo!pliqhasta) & _
             " AND " & buliq_periodo!pliqdesde & "<= elfechahasta " & _
             " AND tipdia.tdconliq = " & NroConce & _
             " ORDER BY emp_lic.elfechadesde "
    OpenRecordset StrSql, rs_Emp_Lic
    
    If Not rs_Emp_Lic.EOF Then
        fecinicio = rs_Emp_Lic!elfechadesde
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Inicio de la licencia " & fecinicio
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontraron Licencias"
        End If
    End If

    If CStr(fecinicio) <> "0:00:00" Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco el periodo Inicial al que pertenece la licencia"
        End If
        StrSql = "SELECT * FROM periodo WHERE periodo.pliqdesde <= " & ConvFecha(fecinicio) & _
                 " AND " & ConvFecha(fecinicio) & " <= periodo.pliqhasta"
        OpenRecordset StrSql, rs_Periodo
    
        If Not rs_Periodo.EOF Then
            nro_desde = rs_Periodo!PliqNro
            nro_hasta = rs_Periodo!PliqNro - 1
            Cant_per = 0
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� el periodo"
            End If
            ' deberia setear lo valores que retorna
            Exit Sub
        End If
    Else
        ' deberia setear lo valores que retorna
        Exit Sub
    End If
    
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco el periodo Final segun la cantidad de Meses a Buscar"
    End If
    StrSql = "SELECT * FROM periodo WHERE pliqnro >= " & nro_desde - Meses & _
             " AND pliqnro <= " & nro_desde & _
             " ORDER BY pliqnro"
    OpenRecordset StrSql, rs_Periodo
    
    ' Seteo la cantidad de periodos encontrados
    Cant_per = rs_Periodo.RecordCount
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & Cant_per & " periodos encontrados"
        Flog.writeline Espacios(Tabulador * 4) & " Busco el acumulador " & NroAcu & " en esos periodos "
    End If
    
    StrSql = "SELECT acu_liq.alcant, acu_liq.almonto FROM periodo " & _
             " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
             " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
             " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
             " WHERE cabliq.empleado = " & NroEmple & "AND periodo.pliqnro >= " & nro_desde - Meses & _
             " AND periodo.pliqnro <= " & nro_desde & _
             " AND acu_liq.acunro = " & NroAcu & _
             " ORDER by periodo.pliqnro"
    OpenRecordset StrSql, RS_Acu
    
    If RS_Acu.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontro ninguna liquidacion con ese acumulador en esos periodos para este empleado " & NroEmple
        End If
    End If
    
    Do While Not RS_Acu.EOF
        If cant Then
            Valor = Valor + RS_Acu!alcant
        Else
            Valor = Valor + RS_Acu!almonto
        End If
    
        RS_Acu.MoveNext
    Loop
    
    Bien = True
    If Cant_per = 0 Then
        Valor = 0
    Else
        Valor = Valor / Cant_per
    End If
    
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If RS_Acu.State = adStateOpen Then RS_Acu.Close
    If rs_Emp_Lic.State = adStateOpen Then rs_Emp_Lic.Close
    
    Set rs_Periodo = Nothing
    Set RS_Acu = Nothing
    Set rs_Emp_Lic = Nothing
End Sub


Public Sub bus_Acum2(NroAcu As Long, Meses As Integer, cant As Boolean, Op As Long, tmes As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca los acumuladores de la licencia con cantidad o monto. gacum2.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

                     
Dim nro_desde As Long
Dim nro_hasta As Long
Dim i As Integer
Dim Cant_per As Single

Dim rs_Periodo As New ADODB.Recordset
Dim RS_Acu As New ADODB.Recordset
    
    Bien = False
    
    If Op = 2 Then
        Valor = -1
    Else
        Valor = 0
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco el periodo anterior "
    End If
    
    ' Encontrar el periodo anterior
    StrSql = "SELECT * FROM periodo WHERE periodo.pliqnro = " & buliq_proceso!PliqNro
    OpenRecordset StrSql, rs_Periodo
    
    If Not rs_Periodo.EOF Then
        nro_desde = rs_Periodo!PliqNro - 1
        nro_hasta = rs_Periodo!PliqNro - 1
        Cant_per = 0
    Else
        ' deberia setear lo valores que retorna
        Bien = False
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� el periodo anterior al periodo " & rs_Periodo!PliqNro
        End If
        Exit Sub
    End If
    
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    
    StrSql = "SELECT * FROM periodo WHERE pliqnro >= " & nro_desde - Meses & _
             " AND pliqnro <= " & nro_desde & _
             " ORDER BY pliqnro"
    OpenRecordset StrSql, rs_Periodo
    
    If Not rs_Periodo.EOF Then
        rs_Periodo.MoveLast
        nro_hasta = rs_Periodo!PliqNro
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Cantidad de periodos encontrados " & rs_Periodo.RecordCount
        Flog.writeline Espacios(Tabulador * 4) & " Busco el acumulador " & NroAcu & " en esos periodos "
    End If
    
    ' Seteo la cantidad de periodos encontrados
    Cant_per = rs_Periodo.RecordCount
    
    StrSql = "SELECT acu_liq.alcant, acu_liq.almonto FROM periodo " & _
             " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
             " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
             " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
             " WHERE cabliq.empleado = " & NroEmple & "AND periodo.pliqnro >= " & nro_desde - Meses & _
             " AND periodo.pliqnro <= " & nro_desde & _
             " AND acu_liq.acunro = " & NroAcu & _
             " ORDER by periodo.pliqnro"
    OpenRecordset StrSql, RS_Acu
    
    If RS_Acu.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontro ninguna liquidacion con ese acumulador en esos periodos para este empleado " & NroEmple
        End If
    End If
    
    If CBool(USA_DEBUG) Then
        Select Case Op
        Case 1:
            Flog.writeline Espacios(Tabulador * 4) & "Operacion: Sumatoria"
        Case 2:
            Flog.writeline Espacios(Tabulador * 4) & "Operacion: Maximo"
        Case 3:
            Flog.writeline Espacios(Tabulador * 4) & "Operacion: Promedio"
        End Select
    End If
    
    Do While Not RS_Acu.EOF
        Select Case Op
        Case 1: ' Sumatoria
            If cant Then
                Valor = Valor + RS_Acu!alcant
            Else
                Valor = Valor + RS_Acu!almonto
            End If
    
        Case 2: ' Maximo
            If cant Then
                If Valor = -1 Then
                    Valor = RS_Acu!alcant
                Else
                    If Valor < RS_Acu!alcant Then
                        Valor = RS_Acu!alcant
                    End If
                End If
            Else
                If Valor = -1 Then
                    Valor = RS_Acu!almonto
                Else
                    If Valor < RS_Acu!almonto Then
                        Valor = RS_Acu!almonto
                    End If
                End If
            End If
            
        Case 3: 'Promedio
            If cant Then
                Valor = RS_Acu!alcant
            Else
                Valor = RS_Acu!almonto
            End If
                      
        Case Else ' no puede ser
        
        End Select
        RS_Acu.MoveNext
    Loop
    
    ' si es promedio ==> lo calculo, sino valor ya trae el valor a retornar
    If Op = 3 Then
        If Cant_per = 0 Then
            Valor = 0
        Else
            Valor = Valor / Cant_per
        End If
    End If
    Bien = True

' Cierro todo y libero
    If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
    If RS_Acu.State = adStateOpen Then RS_Acu.Close
    
    Set rs_Periodo = Nothing
    Set RS_Acu = Nothing
End Sub

Public Sub bus_Acum3()
' ---------------------------------------------------------------------------------------------
' Descripcion: Acum.Mens.Meses Fijos. gacum3.p
'              Obtencion del Acumulador  1 - Semestral    a) Liq Actual
'                                                         b) Liq. Anterior
'
'                                        2 - Anual        a) Liq Actual
'                                                         b) Liq. Anterior
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroAcu As Long          ' Nro de Acumulador
Dim Incluye As Integer      ' 0  - No Incluye
                            ' 1  - Proceso Actual
                            ' 2  - Periodo Actual con Proceso actual
                            ' 3  - Periodo Actual sin proceso actual
                            ' 4  - Anterior
Dim Monto As Boolean        ' True  - MOnto
                            ' False - Cantidad
Dim Opcion As Long          ' 1 - Sumatoria
                            ' 2 - Maximo
                            ' 3 - Promedio
                            ' 4 - Promedio sin 0
                            ' 5 - Minimo
Dim Semestral As Boolean    ' True  = Semestral
                            ' False = Anual
Dim Con_Fases As Boolean    ' True  - Calculo con Fases
                            ' False - Calculo sin Fases
                            
Dim CantMeses As Integer
Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim Cantidad As Single
Dim CantAnios As Integer
Dim BusquedaCompleta As Boolean
Dim DividePor As Integer
Dim UsaActual As Boolean
Dim UsaPeriodoActual As Boolean
'Dim Param_cur As New ADODB.Recordset
    
    
    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        Opcion = Arr_Programa(NroProg).Auxint2
        If Opcion = 3 Then
            If Not EsNulo(Arr_Programa(NroProg).Auxint4) Then
                BusquedaCompleta = IIf(Arr_Programa(NroProg).Auxint4 = 0, True, False)
            Else
                BusquedaCompleta = False
            End If
        End If
        Semestral = CBool(Arr_Programa(NroProg).Auxlog1)
        Incluye = Arr_Programa(NroProg).Auxint3
        'Actual = CBool(Arr_Programa(nroprog).auxlog2)
        Monto = CBool(Arr_Programa(NroProg).Auxlog3)
        Con_Fases = CBool(Arr_Programa(NroProg).Auxlog4)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If


'If Semestral Then
'    If Actual Then
'        If buliq_periodo!pliqmes > 6 Then
'            CantMeses = buliq_periodo!pliqmes - 7
'            MesHasta = buliq_periodo!pliqmes
'            AnioHasta = buliq_periodo!pliqanio
'        Else
'            CantMeses = buliq_periodo!pliqmes - 1
'            MesHasta = buliq_periodo!pliqmes
'            AnioHasta = buliq_periodo!pliqanio
'        End If
'    Else ' Anterior
'        If buliq_periodo!pliqmes > 6 Then
'            CantMeses = 5
'            MesHasta = 6
'            AnioHasta = buliq_periodo!pliqanio
'        Else
'            CantMeses = 7
'            MesHasta = 12
'            AnioHasta = buliq_periodo!pliqanio - 1
'        End If
'    End If
'Else ' Anual
'    If Actual Then
'        CantMeses = buliq_periodo!pliqmes - 1
'        MesHasta = buliq_periodo!pliqmes
'        AnioHasta = buliq_periodo!pliqanio
'    Else
'        CantMeses = 11
'        MesHasta = 12
'        AnioHasta = buliq_periodo!pliqanio - 1
'    End If
'End If

If Semestral Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Semestral "
    End If
    
    Select Case Incluye
    Case 0: 'Semestre actual y no icluye ni periodo ni proceso
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Semestre actual y no icluye ni periodo ni proceso "
        End If
        
        If buliq_periodo!pliqmes > 6 Then
            CantMeses = buliq_periodo!pliqmes - 7
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            CantMeses = buliq_periodo!pliqmes - 1
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
        UsaPeriodoActual = False
    Case 1: 'Semestre actual y Incluye Proceso Actual y no el priodo
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Semestre actual y Incluye Proceso Actual y no el priodo "
        End If
    
        If buliq_periodo!pliqmes > 6 Then
            CantMeses = buliq_periodo!pliqmes - 7
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        Else
            CantMeses = buliq_periodo!pliqmes - 1
            MesHasta = buliq_periodo!pliqmes - 1
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = True
        UsaPeriodoActual = False
    Case 2: 'Semestre actual y Incluye Periodo Actual con proceso actual
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Semestre actual y Incluye Periodo Actual con proceso actual "
        End If
    
        If buliq_periodo!pliqmes > 6 Then
            CantMeses = buliq_periodo!pliqmes - 6
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        Else
            CantMeses = buliq_periodo!pliqmes
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = True
        UsaPeriodoActual = True
    Case 3: 'Semestre actual y Incluye Periodo Actual sin Proceso Actual
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Semestre actual y Incluye Periodo Actual sin Proceso Actual "
        End If
    
        If buliq_periodo!pliqmes > 6 Then
            CantMeses = buliq_periodo!pliqmes - 6
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        Else
            CantMeses = buliq_periodo!pliqmes
            MesHasta = buliq_periodo!pliqmes
            AnioHasta = buliq_periodo!pliqanio
        End If
        UsaActual = False
        UsaPeriodoActual = True
    Case 4: 'Semestre Anterior
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Semestre Anterior "
        End If
    
        If buliq_periodo!pliqmes > 6 Then
            CantMeses = 6
            MesHasta = 6
            AnioHasta = buliq_periodo!pliqanio
        Else
            CantMeses = 6
            MesHasta = 12
            AnioHasta = buliq_periodo!pliqanio - 1
        End If
        UsaActual = False
        UsaPeriodoActual = False
    End Select
    If Opcion = 3 And BusquedaCompleta Then
        DividePor = 6
    Else
        DividePor = CantMeses + 1
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Divide Por " & DividePor
    End If
    
Else ' Anual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Anual "
    End If
    Select Case Incluye
    Case 0: ' A�o Actual y No incluye periodo ni proceso actual
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "A�o Actual y No incluye periodo ni proceso actual "
        End If
    
        CantMeses = buliq_periodo!pliqmes - 1
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
        UsaActual = False
        UsaPeriodoActual = False
    Case 1: 'A�o actual y Incluye Proceso Actual y no el periodo actual
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "A�o actual y Incluye Proceso Actual y no el periodo actual "
        End If
    
        CantMeses = buliq_periodo!pliqmes - 1
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
        UsaActual = True
        UsaPeriodoActual = False
    Case 2: 'A�o actual y Incluye Periodo Actual con proceso actual
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "A�o actual y Incluye Periodo Actual con proceso actual "
        End If
    
        CantMeses = buliq_periodo!pliqmes
        MesHasta = buliq_periodo!pliqmes
        AnioHasta = buliq_periodo!pliqanio
        UsaActual = True
        UsaPeriodoActual = True
    Case 3: 'A�o actual y Incluye Periodo Actual sin proceso actual
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "A�o actual y Incluye Periodo Actual sin proceso actual "
        End If
    
        CantMeses = buliq_periodo!pliqmes
        MesHasta = buliq_periodo!pliqmes
        AnioHasta = buliq_periodo!pliqanio
        UsaActual = False
        UsaPeriodoActual = True
    Case 4: 'A�o Anterior
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "A�o anterior "
        End If
    
        CantMeses = 12
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
        UsaActual = False
        UsaPeriodoActual = False
    End Select
    If Opcion = 3 And BusquedaCompleta Then
        DividePor = 12
    Else
        DividePor = CantMeses + 1
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Divide Por " & DividePor
    End If
End If

Select Case Opcion
Case 1: 'Sumatoria
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Sumatoria "
    End If
    Call AM_Sum(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case 2: 'Maximo
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Maximo "
    End If

    Call AM_Max(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case 3: 'Promedio
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Promedio "
    End If

    Call AM_Prom(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, DividePor, UsaPeriodoActual)
Case 4: 'Promedio sin cero
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Promedio sin cero "
    End If

    Call AM_PromSin0(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case 5: 'Minimo
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Minimo "
    End If

    Call AM_Min(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case Else
End Select


Bien = True
If Not Monto Then
    Valor = Cantidad
End If

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

End Sub

Public Sub bus_Acum4()
' ---------------------------------------------------------------------------------------------
' Descripcion: Acum.Mens.Meses Variables. gacum4.p
'              Obtencion del Acumulador  1 - Semestral    a) Liq Actual
'                                                         b) Liq. Anterior
'
'                                        2 - Anual        a) Liq Actual
'                                                         b) Liq. Anterior
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroAcu As Long          ' Nro de Acumulador
Dim CantMeses As Integer    ' Cantidad de Meses
Dim Opcion As Long          ' 1 - Sumatoria
                            ' 2 - Maximo
                            ' 3 - Promedio
                            ' 4 - Promedio sin 0
                            ' 5 - Minimo
                            
Dim Con_Fases As Boolean     ' True  - Calculo con Fases
                            ' False - Calculo sin Fases
Dim Monto As Boolean        ' True  - MOnto
                            ' False - Cantidad
Dim Incluye As Integer      ' 0  - No Incluye
                            ' 1  - Proceso Actual
                            ' 2  - Periodo Actual sin proceso actual
                            ' 3  - Periodo Actual con Proceso actual
Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim Cantidad As Single
Dim CantAnios As Integer
Dim mesActual1
Dim anioActual1
Dim cantidadDeMeses
Dim K
Dim rsConsult As New ADODB.Recordset
Dim FDesde1
Dim FHasta1

Dim UsaActual As Boolean
Dim UsaPeriodoActual As Boolean
'Dim Param_cur As New ADODB.Recordset
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        CantMeses = Arr_Programa(NroProg).Auxint2
        Opcion = Arr_Programa(NroProg).Auxint3
        Monto = IIf(Arr_Programa(NroProg).Auxint5 = -1 Or Arr_Programa(NroProg).Auxint5 = 2, True, False)
        Incluye = CInt(Arr_Programa(NroProg).Auxint4)
        Con_Fases = CBool(Arr_Programa(NroProg).Auxlog1)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If


Select Case Incluye
Case 0: 'No icluye ni Periodo actual ni proceso actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "No icluye ni Periodo actual ni proceso actual "
    End If

    If buliq_periodo!pliqmes = 1 Then
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
    Else
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
    End If
    UsaActual = False
    UsaPeriodoActual = False
Case 1: ' Incluye Proceso Actual y no periodo actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Proceso Actual y no periodo actual "
    End If

    If buliq_periodo!pliqmes = 1 Then
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
    Else
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
    End If
    UsaActual = True
    UsaPeriodoActual = False
Case 2: 'Incluye Periodo Actual y el Proceso Actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y el Proceso Actual "
    End If

    MesHasta = buliq_periodo!pliqmes
    AnioHasta = buliq_periodo!pliqanio
    UsaActual = True
    UsaPeriodoActual = True
Case 3: 'Incluye Periodo Actual y no el Proceso Actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y no el Proceso Actual "
    End If

    MesHasta = buliq_periodo!pliqmes
    AnioHasta = buliq_periodo!pliqanio
    UsaActual = False
    UsaPeriodoActual = True
End Select

    Bien = False
    Valor = 0

CantAnios = Int(CantMeses / 12)
CantMeses = CantMeses - (CantAnios * 12)


Select Case Opcion
Case 1: 'Sumatoria
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Sumatoria "
    End If

    Call AM_Sum(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case 2: 'Maximo
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Maximo "
    End If

    Call AM_Max(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case 3: 'Promedio
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Promedio "
    End If
    
    'Busco cuantos meses estubo activo el empleado
    cantidadDeMeses = 0
    mesActual1 = MesHasta
    anioActual1 = AnioHasta
    
    For K = 1 To ((CantAnios * 12) + CantMeses)
        
        'Controlo si para el mes actual estubo activo
        FDesde1 = CDate("01/" & mesActual1 & "/" & anioActual1)
        If mesActual1 = 12 Then
           FHasta1 = "01/01/" & (anioActual1 + 1)
        Else
           FHasta1 = "01/" & (mesActual1 + 1) & "/" & anioActual1
        End If
        FHasta1 = DateAdd("d", -1, CDate(FHasta1))
        
        StrSql = " SELECT * FROM fases"
        StrSql = StrSql & " Where real= -1 AND Empleado = " & buliq_empleado!ternro
        StrSql = StrSql & " AND ("
        StrSql = StrSql & "      (altfec <= " & ConvFecha(FHasta1) & " AND bajfec IS NULL)"
        StrSql = StrSql & "  OR  (altfec <= " & ConvFecha(FHasta1) & " AND bajfec >= " & ConvFecha(FDesde1) & ")"
        StrSql = StrSql & "     )"
    
        OpenRecordset StrSql, rsConsult
        
        If Not rsConsult.EOF Then
           cantidadDeMeses = cantidadDeMeses + 1
        End If
        
        rsConsult.Close
        
        mesActual1 = mesActual1 - 1
        If mesActual1 = 0 Then
             mesActual1 = 12
             anioActual1 = anioActual1 - 1
        End If
    Next
    
    Call AM_Prom(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, cantidadDeMeses, UsaPeriodoActual)
'    Call AM_Prom(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, ((CantAnios * 12) + CantMeses), UsaPeriodoActual)
Case 4: 'Promedio sin cero
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Promedio sin cero "
    End If

    Call AM_PromSin0(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case 5: 'Minimo
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Minimo "
    End If

    Call AM_Min(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
Case Else
End Select

If Not Monto Then
    Valor = Cantidad
End If
Bien = True

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

End Sub

Public Sub bus_ImponiblesMensuales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Imponibles Mensuales
'              Obtencion del Acumulador  1 - Semestral    a) Liq Actual
'                                                         b) Liq. Anterior
'
'                                        2 - Anual        a) Liq Actual
'                                                         b) Liq. Anterior
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroAcu As Long           ' Nro de Acumulador
Dim CantMeses As Integer     ' Cantidad de Meses
Dim Opcion As Long           ' 1 - Sumatoria
                             ' 2 - Maximo
                             ' 3 - Promedio
                             ' 4 - Promedio sin 0
                             ' 5 - Minimo
                            
Dim Con_Fases As Boolean     ' True  - Calculo con Fases
                             ' False - Calculo sin Fases
Dim Monto As Boolean         ' True  - MOnto
                             ' False - Cantidad
Dim Incluye As Integer       ' 0  - No Incluye
                             ' 1  - Proceso Actual y no periodo
                             ' 2  - Periodo Actual y no proceso
                             ' 3  - Periodo Actual y proceso

Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim Cantidad As Single
Dim CantAnios As Integer
Dim OpcionValor As String 'Suelgo = A
                            'SAC = B
                            'LAR = C
                            'Todos = T
                            
Dim TipoAcum As Integer     '1 - SUELDO
                            '2 - LAR
                            '3 - SAC
                            '0 - Todos
Dim UsaActual As Boolean
    
    Bien = False
    Valor = 0
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        CantMeses = Arr_Programa(NroProg).Auxint2
        Opcion = Arr_Programa(NroProg).Auxint3
        Monto = Arr_Programa(NroProg).Auxint5
        Incluye = CInt(Arr_Programa(NroProg).Auxint4)
        Con_Fases = CBool(Arr_Programa(NroProg).Auxlog1)
        OpcionValor = Arr_Programa(NroProg).Auxchar1
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Tiene en cuenta las fases "
    Flog.writeline Espacios(Tabulador * 4) & "Tipo de acumulador "
End If

Select Case UCase(OpcionValor)
Case "A":
    TipoAcum = 1
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Sueldo "
    End If
Case "B":
    TipoAcum = 3
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "SAC "
    End If
Case "C":
    TipoAcum = 2
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "LAR "
    End If
Case Else:
    TipoAcum = 0
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "TODOS "
    End If
End Select

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Incluye"
End If
Select Case Incluye
Case 0: 'No icluye ni Periodo actual ni proceso actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "No icluye ni Periodo actual ni proceso actual"
    End If
    If buliq_periodo!pliqmes = 1 Then
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
    Else
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
    End If
    UsaActual = False
Case 1: 'Incluye Proceso Actual y no periodo actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Proceso Actual y no periodo actual"
    End If
    If buliq_periodo!pliqmes = 1 Then
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
    Else
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
    End If
    UsaActual = True
Case 2: 'Incluye Periodo Actual y el Proceso Actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y el Proceso Actual"
    End If
    MesHasta = buliq_periodo!pliqmes
    AnioHasta = buliq_periodo!pliqanio
    UsaActual = True
Case 3: 'Incluye Periodo Actual y no el Proceso Actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y no el Proceso Actual"
    End If
    MesHasta = buliq_periodo!pliqmes
    AnioHasta = buliq_periodo!pliqanio
    UsaActual = False
End Select



CantAnios = Int(CantMeses / 12)
CantMeses = CantMeses - (CantAnios * 12)

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Operacion: "
End If
Select Case Opcion
Case 1: 'Sumatoria
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Sumatoria"
    End If
    Call AM_Sum_Nuevo(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, True, UsaActual, TipoAcum)
Case 2: 'Maximo
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Maximo"
    End If
    Call AM_Max_Nuevo(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, True, UsaActual, TipoAcum)
Case 3: 'Promedio
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Promedio"
    End If
    Call AM_Prom_Nuevo(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, True, UsaActual, TipoAcum, (CantAnios * 12) + CantMeses)
Case 4: 'Promedio sin cero
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Promedio sin cero"
    End If
    Call AM_PromSin0_Nuevo(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, True, UsaActual, TipoAcum)
Case 5: 'Minimo
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Minimo"
    End If
    Call AM_Min_Nuevo(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, True, UsaActual, TipoAcum)
Case Else
End Select

If Not Monto Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Retorna cantidad"
    End If
    Valor = Cantidad
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Retorna monto"
    End If
End If
End Sub

Public Sub bus_ImponiblesDelProceso()
' ---------------------------------------------------------------------------------------------
' Descripcion: Acumuladores Imponibles del Proceso
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroAcu As Long           ' Nro de Acumulador

Dim Monto As Integer         ' 1 - Cantidad
                             ' 2 - Monto
                             
                             
Dim Obligatorio As Boolean   'Obligatorio retornar valor

Dim TipoAcu As Integer     '1 - SUELDO
                            '2 - SAC
                            '3 - LAR
                            '4 - Todos

'Dim Param_cur As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset
    
    Valor = 0
    Bien = False
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        Monto = Arr_Programa(NroProg).Auxint2
        Obligatorio = CBool(Arr_Programa(NroProg).Auxlog1)
        TipoAcu = Arr_Programa(NroProg).Auxint3
    Else
        Exit Sub
    End If

'busco el acu_liq de este proceso
If TipoAcu = 4 Then 'Todos los tipos
    StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
             " AND acunro =" & NroAcu
Else
    StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
             " AND acunro =" & NroAcu & " AND tconnro = " & TipoAcu
End If
OpenRecordset StrSql, rs_ImpPro
If Not rs_ImpPro.EOF Then
    If Monto = 1 Then
        Valor = rs_ImpPro!ipacant
    Else
        Valor = rs_ImpPro!ipamonto
    End If
    Bien = True
Else
    If Obligatorio Then
        Bien = False
    Else
        Bien = True
    End If
End If

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_ImpPro.State = adStateOpen Then rs_ImpPro.Close
Set rs_ImpPro = Nothing

End Sub


Public Sub AM_Sum(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal UsaPeriodoActual As Boolean)
' sumariza los meses anteriores, INCLUYENDO ACTUAL.
' Parametros : 1) Nro Acumulador                            -  Acu
'              2) Mes  hasta (mes en el que estoy parado)   -
'              3) Anio hasta (anio que estoy parado)        -
'              4) Cant meses anteriores                     -
'              5) Cant Anios anteriores                     -
'              6) con-fases                                 -
'              7) cantidad                                  -
'              8) monto.                                    -


Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios

If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If

' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 4) & "la empfaltagr del empleado es > a " & CDate("01/" & MesDesde & "/" & AnioDesde)
'        End If
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Busca desde el mes " & MesDesde & " del a�o " & AnioDesde & " hasta el mes " & MesHasta & " del a�o " & AnioHasta
End If

If Not Imponibles Then
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesDesde
        Else
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta
        End If
    Else
'        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'                 " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        Else
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        End If
    Else
'        StrSql = "SELECT * FROM acu_mes " & _
'                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acu_mes.acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'                 " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
'                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        StrSql = "SELECT * FROM acu_mes " & _
                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acu_mes.acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
    End If
End If
StrSql = StrSql & " ORDER BY amanio, ammes"
OpenRecordset StrSql, rs_Acu_Mes

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "SQL " & StrSql
End If
If Not rs_Acu_Mes.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Acumulando ..."
    End If
End If

Do While Not rs_Acu_Mes.EOF
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Suma Monto " & IIf(Not EsNulo(rs_Acu_Mes!ammonto), rs_Acu_Mes!ammonto, 0)
            Flog.writeline Espacios(Tabulador * 4) & "Suma Cantidad " & IIf(Not EsNulo(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
        End If
        
        Valor = Valor + IIf(Not EsNulo(rs_Acu_Mes!ammonto), rs_Acu_Mes!ammonto, 0)
        Cantidad = Cantidad + IIf(Not EsNulo(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
    rs_Acu_Mes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
'If (MesHasta = buliq_periodo!pliqmes And AnioHasta = buliq_periodo!pliqanio) Or UsaActual Then
If UsaActual Then
    ' FGZ - 09/02/2004
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acu)) Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Suma Proceso Actual: "
                Flog.writeline Espacios(Tabulador * 5) & "Monto " & objCache_Acu_Liq_Monto.Valor(CStr(Acu))
                Flog.writeline Espacios(Tabulador * 6) & "Cantidad " & objCache_Acu_Liq_Cantidad.Valor(CStr(Acu))
            End If
    
            Valor = Valor + objCache_Acu_Liq_Monto.Valor(CStr(Acu))
            Cantidad = Cantidad + objCache_Acu_Liq_Cantidad.Valor(CStr(Acu))
    End If

'    StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro & _
'             " AND acunro =" & Acu
'    OpenRecordset StrSql, rs_Acu_Liq
'    If Not rs_Acu_Liq.EOF Then
'        Valor = Valor + rs_Acu_Liq!almonto
'        Cantidad = Cantidad + rs_Acu_Liq!alcant
'    End If
End If

' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
'If rs_Acu_liq.State = adStateOpen Then rs_Acu_Liq.Close
Set rs_Fases = Nothing
Set rs_Acu_Mes = Nothing
'Set rs_Acu_Liq = Nothing

End Sub

Public Sub AM_Sum_Nuevo(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal TipoAcu As Integer)
' sumariza los meses anteriores, INCLUYENDO ACTUAL.
' Parametros : 1) Nro Acumulador                            -  Acu
'              2) Mes  hasta (mes en el que estoy parado)   -
'              3) Anio hasta (anio que estoy parado)        -
'              4) Cant meses anteriores                     -
'              5) Cant Anios anteriores                     -
'              6) con-fases                                 -
'              7) cantidad                                  -
'              8) monto.                                    -


Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_ImpMes As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios


If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If TipoAcu = 0 Then 'Todos los tipos
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND impmesarg.tconnro =" & TipoAcu & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
End If
StrSql = StrSql & " ORDER BY imaanio, imames"
OpenRecordset StrSql, rs_ImpMes
If Not rs_ImpMes.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Datos encontrados "
    End If
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "No se encontraron Datos "
        Flog.writeline Espacios(Tabulador * 5) & "SQL:  " & StrSql
    End If
End If

Do While Not rs_ImpMes.EOF
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 6) & "Monto " & rs_ImpMes!imamonto & " Cantidad " & rs_ImpMes!imacant
    End If
    Valor = Valor + IIf(Not EsNulo(rs_ImpMes!imamonto), rs_ImpMes!imamonto, 0)
    Cantidad = Cantidad + IIf(Not EsNulo(rs_ImpMes!imacant), rs_ImpMes!imacant, 0)
   
    rs_ImpMes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
'If (MesHasta = buliq_periodo!pliqmes And AnioHasta = buliq_periodo!pliqanio) Or UsaActual Then
If UsaActual Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Actual "
    End If

    If TipoAcu = 0 Then 'Todos los tipos
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu
    Else
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu & " AND tconnro = " & TipoAcu
    End If
    OpenRecordset StrSql, rs_ImpPro
    If Not rs_ImpPro.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 6) & "Monto " & IIf(Not EsNulo(rs_ImpPro!ipamonto), rs_ImpPro!ipamonto, 0)
            Flog.writeline Espacios(Tabulador * 6) & "Cantidad " & IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
        End If
        Valor = Valor + IIf(Not EsNulo(rs_ImpPro!ipamonto), rs_ImpPro!ipamonto, 0)
        Cantidad = Cantidad + IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
    End If
End If

' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_ImpPro.State = adStateOpen Then rs_ImpPro.Close
If rs_ImpMes.State = adStateOpen Then rs_ImpMes.Close
Set rs_Fases = Nothing
Set rs_ImpMes = Nothing
Set rs_ImpPro = Nothing

End Sub

Public Sub AM_Max_Nuevo(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal TipoAcu As Integer)
' sumariza los meses anteriores, INCLUYENDO ACTUAL.
' Parametros : 1) Nro Acumulador                            -  Acu
'              2) Mes  hasta (mes en el que estoy parado)   -
'              3) Anio hasta (anio que estoy parado)        -
'              4) Cant meses anteriores                     -
'              5) Cant Anios anteriores                     -
'              6) con-fases                                 -
'              7) cantidad                                  -
'              8) monto.                                    -


Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_ImpMes As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset

Valor = 0
Cantidad = 0

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios

If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If
If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 5) & "Mes desde " & MesDesde & " a�o desde " & AnioDesde
    Flog.writeline Espacios(Tabulador * 5) & "Mes Hasta " & MesHasta & " a�o hasta " & AnioHasta
End If

' Modificado para que tome el promedio para los jornales
If ConFases Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Busco la ultima fase activa"
    End If
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If CBool(USA_DEBUG) Then
            Flog.Write Espacios(Tabulador * 5) & "Fase desde " & rs_Fases!altfec
            If Not EsNulo(rs_Fases!bajfec) Then
                Flog.writeline Espacios(Tabulador * 5) & " hasta " & rs_Fases!bajfec
            Else
                Flog.writeline Espacios(Tabulador * 5) & " hasta NULL "
            End If
        End If
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 5) & "No se encontr� fase activa"
        End If
    End If
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Emleado empfaltagr " & buliq_empleado!empfaltagr
    End If
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If TipoAcu = 0 Then 'Todos los tipos
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND impmesarg.tconnro =" & TipoAcu & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
End If
StrSql = StrSql & " ORDER BY imaanio, imames"
OpenRecordset StrSql, rs_ImpMes

'' FGZ - 18/04/2004
'If TipoAcu = 0 Then 'Todos los tipos
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'Else
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND impmesarg.tconnro =" & TipoAcu & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'End If
'StrSql = StrSql & " ORDER BY imaanio, imames"
'OpenRecordset StrSql, rs_ImpMes
If Not rs_ImpMes.EOF Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "Datos encontrados "
    End If
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 5) & "No se encontraron Datos "
        Flog.writeline Espacios(Tabulador * 5) & "SQL:  " & StrSql
    End If
End If

Do While Not rs_ImpMes.EOF
    If Not EsNulo(rs_ImpMes!imamonto) Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 6) & "Monto " & rs_ImpMes!imamonto
        End If
        If rs_ImpMes!imamonto > Valor Then
            Valor = rs_ImpMes!imamonto
            Cantidad = IIf(Not EsNulo(rs_ImpMes!imacant), rs_ImpMes!imacant, 0)
        End If
    End If
    rs_ImpMes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    If TipoAcu = 0 Then 'Todos los tipos
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu
    Else
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu & " AND tconnro = " & TipoAcu
    End If
    OpenRecordset StrSql, rs_ImpPro
    If Not rs_ImpPro.EOF Then
        If Not EsNulo(rs_ImpPro!ipamonto) Then
            If rs_ImpPro!ipamonto > Valor Then
                Valor = rs_ImpPro!ipamonto
                Cantidad = IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
            End If
        End If
    End If
End If

' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_ImpMes.State = adStateOpen Then rs_ImpMes.Close
If rs_ImpPro.State = adStateOpen Then rs_ImpPro.Close
Set rs_Fases = Nothing
Set rs_ImpMes = Nothing
Set rs_ImpPro = Nothing

End Sub

Public Sub AM_Max(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal UsaPeriodoActual As Boolean)
' sumariza los meses anteriores, INCLUYENDO ACTUAL.
' Parametros : 1) Nro Acumulador                            -  Acu
'              2) Mes  hasta (mes en el que estoy parado)   -
'              3) Anio hasta (anio que estoy parado)        -
'              4) Cant meses anteriores                     -
'              5) Cant Anios anteriores                     -
'              6) con-fases                                 -
'              7) cantidad                                  -
'              8) monto.                                    -


Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset
Dim Aux_Valor As Single
Dim Aux_Cant As Single

Valor = 0
Cantidad = 0
Aux_Valor = 0
Aux_Cant = 0

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios

If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If Not Imponibles Then
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesDesde
        Else
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        Else
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        End If
    Else
        StrSql = "SELECT * FROM acu_mes " & _
                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acu_mes.acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
    End If
End If
StrSql = StrSql & " ORDER BY amanio, ammes"
OpenRecordset StrSql, rs_Acu_Mes

'    If Not Imponibles Then
'        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'                 " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
'    Else
'        StrSql = "SELECT * FROM acu_mes " & _
'                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acu_mes.acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'                 " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
'                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
'    End If
'StrSql = StrSql & " ORDER BY amanio, ammes"
'OpenRecordset StrSql, rs_Acu_Mes

Do While Not rs_Acu_Mes.EOF
    If Not EsNulo(rs_Acu_Mes!ammonto) Then
        If UsaPeriodoActual Then
            If rs_Acu_Mes!ammes = buliq_periodo!pliqmes Then
                Aux_Valor = Aux_Valor + rs_Acu_Mes!ammonto
                Aux_Cant = Aux_Cant + rs_Acu_Mes!amcant
            End If
        End If
        If rs_Acu_Mes!ammonto > Valor Then
            Valor = rs_Acu_Mes!ammonto
            Cantidad = IIf(Not EsNulo(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
        End If
    End If
    rs_Acu_Mes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    ' FGZ - 09/02/2004
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acu)) Then
        If objCache_Acu_Liq_Monto.Valor(CStr(Acu)) + Aux_Valor > Valor Then
            Valor = objCache_Acu_Liq_Monto.Valor(CStr(Acu)) + Aux_Valor
            Cantidad = objCache_Acu_Liq_Cantidad.Valor(CStr(Acu)) + Aux_Cant
        End If
    End If
End If
'FGZ - 09/02/2004
' esto se saca porque ahora estan en el cahce de acu_liq

'    StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro & _
'             " AND acunro =" & Acu
'    OpenRecordset StrSql, rs_Acu_Liq
'    If Not rs_Acu_Liq.EOF Then
'        If rs_Acu_Liq!ammonto > Valor Then
'            Valor = rs_Acu_Liq!ammonto
'            Cantidad = rs_Acu_Liq!amcant
'        End If
'    End If
'End If

' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
'If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
Set rs_Fases = Nothing
Set rs_Acu_Mes = Nothing
'Set rs_Acu_Liq = Nothing

End Sub


Public Sub AM_Min(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal UsaPeriodoActual As Boolean)
' Busca en los meses anteriores, INCLUYENDO ACTUAL.
' Parametros : 1) Nro Acumulador                            -  Acu
'              2) Mes  hasta (mes en el que estoy parado)   -
'              3) Anio hasta (anio que estoy parado)        -
'              4) Cant meses anteriores                     -
'              5) Cant Anios anteriores                     -
'              6) con-fases                                 -
'              7) cantidad                                  -
'              8) monto.                                    -


Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset

Dim Encontro As Boolean
Dim Aux_Valor As Single
Dim Aux_Cant As Single


Encontro = False

Valor = 0
Cantidad = 0
Aux_Valor = 0
Aux_Cant = 0


If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios

If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If Not Imponibles Then
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesDesde
        Else
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        Else
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        End If
    Else
        StrSql = "SELECT * FROM acu_mes " & _
                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acu_mes.acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
    End If
End If
StrSql = StrSql & " ORDER BY amanio, ammes"
OpenRecordset StrSql, rs_Acu_Mes

'If Not Imponibles Then
'    StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
'             " AND acunro =" & Acu & _
'             " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'             " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
'Else
'    StrSql = "SELECT * FROM acu_mes " & _
'             " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
'             " WHERE ternro = " & buliq_empleado!ternro & _
'             " AND acu_mes.acunro =" & Acu & _
'             " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'             " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
'             " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
'End If
'StrSql = StrSql & " ORDER BY amanio, ammes"
'OpenRecordset StrSql, rs_Acu_Mes

If Not rs_Acu_Mes.EOF Then
    'Me muevo al primero
    rs_Acu_Mes.MoveFirst
    If Not EsNulo(rs_Acu_Mes!ammonto) Then
        Encontro = True
        Valor = rs_Acu_Mes!ammonto
        Cantidad = rs_Acu_Mes!amcant
    End If
End If
rs_Acu_Mes.MoveNext

Do While Not rs_Acu_Mes.EOF
    If Not EsNulo(rs_Acu_Mes!ammonto) Then
        If UsaPeriodoActual Then
            If rs_Acu_Mes!ammes = buliq_periodo!pliqmes Then
                Aux_Valor = Aux_Valor + rs_Acu_Mes!ammonto
                Aux_Cant = Aux_Cant + rs_Acu_Mes!amcant
            End If
        End If
        If rs_Acu_Mes!ammonto < Valor Then
            Valor = rs_Acu_Mes!ammonto
            Cantidad = IIf(Not EsNulo(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
        End If
    End If
    rs_Acu_Mes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    ' FGZ - 09/02/2004
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acu)) Then
        If Not Encontro Then
            Valor = objCache_Acu_Liq_Monto.Valor(CStr(Acu))
            Cantidad = objCache_Acu_Liq_Cantidad.Valor(CStr(Acu))
        Else
            If objCache_Acu_Liq_Monto.Valor(CStr(Acu)) + Aux_Valor < Valor Then
                Valor = objCache_Acu_Liq_Monto.Valor(CStr(Acu)) + Aux_Valor
                Cantidad = objCache_Acu_Liq_Cantidad.Valor(CStr(Acu)) + Aux_Cant
            End If
        End If
    End If
End If
    
' FGZ - 09/02/2004
' esto se saco porque ahora esta en el cache de acu_liq

'    StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro & _
'             " AND acunro =" & Acu
'    OpenRecordset StrSql, rs_Acu_Liq
'    If Not rs_Acu_Liq.EOF Then
'        If Not Encontro Then ' valor todavia no se seteo ==> retorno el del acu_liq de este proceso
'            Valor = rs_Acu_Liq!ammonto
'            Cantidad = rs_Acu_Liq!amcant
'        Else
'            If rs_Acu_Liq!ammonto < Valor Then
'                Valor = rs_Acu_Liq!ammonto
'                Cantidad = rs_Acu_Liq!amcant
'            End If
'        End If
'    End If
'End If

' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
'If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
Set rs_Fases = Nothing
Set rs_Acu_Mes = Nothing
'Set rs_Acu_Liq = Nothing

End Sub


Public Sub AM_Min_Nuevo(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal TipoAcu As Integer)
' Busca en los meses anteriores, INCLUYENDO ACTUAL.
' Parametros : 1) Nro Acumulador                            -  Acu
'              2) Mes  hasta (mes en el que estoy parado)   -
'              3) Anio hasta (anio que estoy parado)        -
'              4) Cant meses anteriores                     -
'              5) Cant Anios anteriores                     -
'              6) con-fases                                 -
'              7) cantidad                                  -
'              8) monto.                                    -


Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_ImpMes As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset

Dim Encontro As Boolean

Encontro = False

Valor = 0
Cantidad = 0

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios

If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If TipoAcu = 0 Then 'Todos los tipos
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND impmesarg.tconnro =" & TipoAcu & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
End If
StrSql = StrSql & " ORDER BY imaanio, imames"
OpenRecordset StrSql, rs_ImpMes

'If TipoAcu = 0 Then 'Todos los tipos
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'Else
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND impmesarg.tconnro =" & TipoAcu & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'End If
'StrSql = StrSql & " ORDER BY imaanio, imames"
'OpenRecordset StrSql, rs_ImpMes

If Not rs_ImpMes.EOF Then
    'Me muevo al primero
    rs_ImpMes.MoveFirst
    If Not EsNulo(rs_ImpMes!imamonto) Then
        Encontro = True
        Valor = rs_ImpMes!imamonto
        Cantidad = IIf(Not EsNulo(rs_ImpMes!imacant), rs_ImpMes!imacant, 0)
    End If
End If
rs_ImpMes.MoveNext

Do While Not rs_ImpMes.EOF
        If rs_ImpMes!imamonto < Valor Then
            Valor = rs_ImpMes!imamonto
            Cantidad = rs_ImpMes!imacant
        End If
    rs_ImpMes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    If TipoAcu = 0 Then 'Todos los tipos
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu
    Else
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu & " AND tconnro = " & TipoAcu
    End If
    OpenRecordset StrSql, rs_ImpPro
    If Not rs_ImpPro.EOF Then
        If Not Encontro Then ' valor todavia no se seteo ==> retorno el del acu_liq de este proceso
            Valor = IIf(Not EsNulo(rs_ImpPro!ipamonto), rs_ImpPro!ipamonto, 0)
            Cantidad = IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
        Else
            If Not EsNulo(rs_ImpPro!ipamonto) Then
                If rs_ImpPro!ipamonto < Valor Then
                    Valor = rs_ImpPro!ipamonto
                    Cantidad = IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
                End If
            End If
        End If
    End If
End If

' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_ImpMes.State = adStateOpen Then rs_ImpMes.Close
If rs_ImpPro.State = adStateOpen Then rs_ImpPro.Close
Set rs_Fases = Nothing
Set rs_ImpPro = Nothing
Set rs_ImpMes = Nothing

End Sub


Public Sub bus_Acum5()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del promedio del acumulador para asignaciones familiares. gacum5.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: FGZ - 02/04/2004
' Descripcion: se agreg� un parametro. Promedio con o sin Cero.
'              Si este parametro es nulo utiliza promedio sin 0 y sino
'              se evalua y se usa la funcion que corresponda
' ---------------------------------------------------------------------------------------------

Dim NroAcu As Long          ' Nro de Acumulador - acumulador.acunro
Dim Con_Fases As Boolean    ' True  - Calculo con fases

Dim CantMeses As Integer
Dim CantAnios As Integer
Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim Aux_Anio As Integer
Dim Cantidad As Single
Dim Con_cero As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset
Dim RS_Acu As New ADODB.Recordset
    
    Bien = False
    Valor = 0

    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        Con_Fases = CBool(Arr_Programa(NroProg).Auxlog1)
        Con_cero = Not IIf(Not EsNulo(Arr_Programa(NroProg).Auxlog2), CBool(Arr_Programa(NroProg).Auxlog2), True)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda "
        End If
        Exit Sub
    End If


    If buliq_periodo!pliqmes > 8 Then
        Aux_Anio = buliq_periodo!pliqanio
    Else
        Aux_Anio = buliq_periodo!pliqanio - 1
    End If
    
    CantMeses = 6
    
    If buliq_periodo!pliqmes >= 3 And buliq_periodo!pliqmes <= 8 Then
        MesHasta = 12
    Else
        MesHasta = 6
    End If
    AnioHasta = Aux_Anio
    
    ' Promedio de los meses anteriores
    If Con_cero Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Promedio con ceros de los meses anteriores "
        End If
        Call AM_Prom(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, False, CantMeses, False)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Promedio sin ceros de los meses anteriores "
        End If
        Call AM_PromSin0(NroAcu, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, False, False)
    End If

    If Valor = 0 Then
        ' FGZ - 09/02/2004
        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(NroAcu)) Then
            Valor = objCache_Acu_Liq_Monto.Valor(CStr(NroAcu))
            Cantidad = objCache_Acu_Liq_Cantidad.Valor(CStr(NroAcu))
        End If
        
'        StrSql = "SELECT * FROM acu_liq " & _
'                 " AND acu_liq.cliqnro = " & buliq_cabliq!cliqnro & _
'                 " AND acu_liq.acunro = " & NroAcu
'        OpenRecordset StrSql, RS_Acu
'
'        If Not RS_Acu.EOF Then
'            Valor = RS_Acu!almonto
'            Cantidad = RS_Acu!alcant
'        End If
    End If
    
    Bien = True
    
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If RS_Acu.State = adStateOpen Then RS_Acu.Close
'Set Param_cur = Nothing
Set RS_Acu = Nothing

End Sub

Public Sub AM_Prom(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal DividePor As Integer, ByVal UsaPeriodoActual As Boolean)
' Promedio de los meses anteriores, incluyendo el actual, toma en cuenta
' la primera fase del empleado o la fecha de alta del empleado
' -------------------------------------------------------------
'Parametros
    '1 - Nro de acumulador                          - Acu
    '2 - Mes Hasta, el mes en el que estoy parado   - MesHasta
    '3 - A�o hasta, a�o en el que estoy parado      - AnioHasta
    '4 - Cantidad de Meses anteiores                - CantMeses
    '5 - Cantidad de A�os anteriores                - CantAnios
    '6 - Con fases                                  - ConFases
    '7 - Monto                                      - ByRef Valor
    '8 - Cantidad                                   - Byref Cantidad
    '10 - Por la cantidad de Meses que divide
' ------------------------------------------------------------

Dim i As Integer
Dim j As Integer
Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios


If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If Not Imponibles Then
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesDesde
        Else
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        Else
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        End If
    Else
        StrSql = "SELECT * FROM acu_mes " & _
                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acu_mes.acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
    End If
End If
StrSql = StrSql & " ORDER BY amanio, ammes"
OpenRecordset StrSql, rs_Acu_Mes

'If Not Imponibles Then
'    StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
'             " AND acunro =" & Acu & _
'             " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'             " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
'Else
'    StrSql = "SELECT * FROM acu_mes " & _
'             " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
'             " WHERE ternro = " & buliq_empleado!ternro & _
'             " AND acu_mes.acunro =" & Acu & _
'             " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'             " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
'             " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
'End If
'StrSql = StrSql & " ORDER BY amanio, ammes"
'OpenRecordset StrSql, rs_Acu_Mes

cantProm = 0
Do While Not rs_Acu_Mes.EOF
        Valor = Valor + IIf(Not EsNulo(rs_Acu_Mes!ammonto), rs_Acu_Mes!ammonto, 0)
        Cantidad = Cantidad + IIf(Not EsNulo(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
        cantProm = cantProm + 1
    rs_Acu_Mes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    ' FGZ - 09/02/2004
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acu)) Then
        Valor = Valor + objCache_Acu_Liq_Monto.Valor(CStr(Acu))
        Cantidad = Cantidad + objCache_Acu_Liq_Cantidad.Valor(CStr(Acu))
        cantProm = cantProm + 1
    End If

'    StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro & _
'             " AND acunro =" & Acu
'    OpenRecordset StrSql, rs_Acu_Liq
'    If Not rs_Acu_Liq.EOF Then
'        Valor = Valor + rs_Acu_Liq!almonto
'        Cantidad = Cantidad + rs_Acu_Liq!alcant
'        cantProm = cantProm + 1
'    End If
End If

If cantProm <> 0 Then
    'Valor = Valor / cantProm
    Valor = Valor / DividePor
    'Cantidad = Cantidad / cantProm
    Cantidad = Cantidad / DividePor
Else
    Valor = 0
    Cantidad = 0
End If


' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
'If rs_Acu_liq.State = adStateOpen Then rs_Acu_Liq.Close
Set rs_Fases = Nothing
Set rs_Acu_Mes = Nothing
'Set rs_Acu_Liq = Nothing

End Sub

Public Sub AM_Prom_Nuevo(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal TipoAcu As Integer, ByVal DividePor As Integer)
' Promedio de los meses anteriores, incluyendo el actual, toma en cuenta
' la primera fase del empleado o la fecha de alta del empleado
' -------------------------------------------------------------
'Parametros
    '1 - Nro de acumulador                          - Acu
    '2 - Mes Hasta, el mes en el que estoy parado   - MesHasta
    '3 - A�o hasta, a�o en el que estoy parado      - AnioHasta
    '4 - Cantidad de Meses anteiores                - CantMeses
    '5 - Cantidad de A�os anteriores                - CantAnios
    '6 - Con fases                                  - ConFases
    '7 - Monto                                      - ByRef Valor
    '8 - Cantidad                                   - Byref Cantidad
' ------------------------------------------------------------

Dim i As Integer
Dim j As Integer
Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_ImpMes As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios


If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If TipoAcu = 0 Then 'Todos los tipos
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND impmesarg.tconnro =" & TipoAcu & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
End If
StrSql = StrSql & " ORDER BY imaanio, imames"
OpenRecordset StrSql, rs_ImpMes


'If TipoAcu = 0 Then 'Todos los tipos
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'Else
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND impmesarg.tconnro =" & TipoAcu & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'End If
'StrSql = StrSql & " ORDER BY imaanio, imames"
'OpenRecordset StrSql, rs_ImpMes

cantProm = 0
Do While Not rs_ImpMes.EOF
    Valor = Valor + IIf(Not EsNulo(rs_ImpMes!imamonto), rs_ImpMes!imamonto, 0)
    Cantidad = Cantidad + IIf(Not EsNulo(rs_ImpMes!imacant), rs_ImpMes!imacant, 0)
    cantProm = cantProm + 1
    
    rs_ImpMes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    If TipoAcu = 0 Then 'Todos los tipos
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu
    Else
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu & " AND tconnro = " & TipoAcu
    End If
    OpenRecordset StrSql, rs_ImpPro
    If Not rs_ImpPro.EOF Then
        If Not EsNulo(rs_ImpPro!ipamonto) Then
            Valor = Valor + rs_ImpPro!ipamonto
            Cantidad = Cantidad + IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
            cantProm = cantProm + 1
        End If
    End If
End If

If cantProm <> 0 Then
    'Valor = Valor / cantProm
    Valor = Valor / DividePor
    'Cantidad = Cantidad / cantProm
    Cantidad = Cantidad / DividePor
Else
    Valor = 0
    Cantidad = 0
End If


' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_ImpPro.State = adStateOpen Then rs_ImpPro.Close
If rs_ImpMes.State = adStateOpen Then rs_ImpMes.Close
Set rs_Fases = Nothing
Set rs_ImpPro = Nothing
Set rs_ImpMes = Nothing

End Sub

Public Sub AM_PromSin0(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal UsaPeriodoActual As Boolean)
' Promedio de los meses anteriores, incluyendo el actual, toma en cuenta
' la primera fase del empleado o la fecha de alta del empleado
' Sin tener en cuenta los que son 0
' -------------------------------------------------------------
'Parametros
    '1 - Nro de acumulador                          - Acu
    '2 - Mes Hasta, el mes en el que estoy parado   - MesHasta
    '3 - A�o hasta, a�o en el que estoy parado      - AnioHasta
    '4 - Cantidad de Meses anteiores                - CantMeses
    '5 - Cantidad de A�os anteriores                - CantAnios
    '6 - Con fases                                  - ConFases
    '7 - Monto                                      - ByRef Valor
    '8 - Cantidad                                   - Byref Cantidad
' ------------------------------------------------------------

Dim i As Integer
Dim j As Integer
Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Acu_Liq As New ADODB.Recordset

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios

If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If Not Imponibles Then
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesDesde
        Else
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes =" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        Else
            StrSql = "SELECT * FROM acu_mes " & _
                     " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acu_mes.acunro =" & Acu & _
                     " AND " & AnioDesde & " = amanio " & _
                     " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta & _
                     " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
        End If
    Else
        StrSql = "SELECT * FROM acu_mes " & _
                 " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acu_mes.acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
                 " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
                 " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
                 " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
    End If
End If
StrSql = StrSql & " ORDER BY amanio, ammes"
OpenRecordset StrSql, rs_Acu_Mes

'If Not Imponibles Then
'    StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
'             " AND acunro =" & Acu & _
'             " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'             " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
'Else
'    StrSql = "SELECT * FROM acu_mes " & _
'             " INNER JOIN acumulador ON acu_mes.acunro = acumulador.acunro " & _
'             " WHERE ternro = " & buliq_empleado!ternro & _
'             " AND acu_mes.acunro =" & Acu & _
'             " AND " & AnioDesde & " <= amanio AND amanio <= " & AnioHasta & _
'             " AND (( ammes >=" & MesDesde & " AND " & AnioDesde & " <= amanio) OR ( ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))" & _
'             " AND (acumulador.acuimponible = -1 or acuimpcont = -1)"
'End If
'StrSql = StrSql & " ORDER BY amanio, ammes"
'OpenRecordset StrSql, rs_Acu_Mes

cantProm = 0
Do While Not rs_Acu_Mes.EOF
    If Not EsNulo(rs_Acu_Mes!ammonto) Then
        If rs_Acu_Mes!ammonto <> 0 Then ' si el monto es cero ==> no lo tengo en cuenta
            Valor = Valor + rs_Acu_Mes!ammonto
            Cantidad = Cantidad + IIf(Not EsNulo(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
            cantProm = cantProm + 1
        End If
    End If
    rs_Acu_Mes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    ' FGZ - 09/02/2004
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(Acu)) Then
        If objCache_Acu_Liq_Monto.Valor(CStr(Acu)) <> 0 Then
            Valor = Valor + objCache_Acu_Liq_Monto.Valor(CStr(Acu))
            Cantidad = Cantidad + objCache_Acu_Liq_Cantidad.Valor(CStr(Acu))
            cantProm = cantProm + 1
        End If
    End If

'    StrSql = "SELECT * FROM acu_liq WHERE cliqnro = " & buliq_cabliq!cliqnro & _
'             " AND acunro =" & Acu
'    OpenRecordset StrSql, rs_Acu_Liq
'    If Not rs_Acu_Liq.EOF Then
'        If rs_Acu_Liq!almonto <> 0 Then ' si el monto es cero ==> no lo tengo en cuenta
'            Valor = Valor + rs_Acu_Liq!almonto
'            Cantidad = Cantidad + rs_Acu_Liq!alcant
'            cantProm = cantProm + 1
'        End If
'    End If
End If

If cantProm <> 0 Then
    Valor = Valor / cantProm
    Cantidad = Cantidad / cantProm
Else
    Valor = 0
    Cantidad = 0
End If


' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
'If rs_Acu_Liq.State = adStateOpen Then rs_Acu_Liq.Close
Set rs_Fases = Nothing
Set rs_Acu_Mes = Nothing
'Set rs_Acu_Liq = Nothing

End Sub

Public Sub AM_PromSin0_Nuevo(ByVal Acu As Long, ByVal MesHasta As Integer, ByVal AnioHasta As Integer, ByVal CantMeses As Integer, ByVal CantAnios As Integer, ByVal ConFases As Boolean, ByRef Valor As Single, ByRef Cantidad As Single, ByVal Imponibles As Boolean, ByVal UsaActual As Boolean, ByVal TipoAcu As Integer)
' Promedio de los meses anteriores, incluyendo el actual, toma en cuenta
' la primera fase del empleado o la fecha de alta del empleado
' Sin tener en cuenta los que son 0
' -------------------------------------------------------------
'Parametros
    '1 - Nro de acumulador                          - Acu
    '2 - Mes Hasta, el mes en el que estoy parado   - MesHasta
    '3 - A�o hasta, a�o en el que estoy parado      - AnioHasta
    '4 - Cantidad de Meses anteiores                - CantMeses
    '5 - Cantidad de A�os anteriores                - CantAnios
    '6 - Con fases                                  - ConFases
    '7 - Monto                                      - ByRef Valor
    '8 - Cantidad                                   - Byref Cantidad
' ------------------------------------------------------------

Dim i As Integer
Dim j As Integer
Dim Hasta As Integer
Dim cantProm As Integer
Dim MesDesde As Integer
Dim AnioDesde As Integer

Dim rs_Fases As New ADODB.Recordset
Dim rs_ImpMes As New ADODB.Recordset
Dim rs_ImpPro As New ADODB.Recordset

If CantMeses > 12 Or MesHasta > 12 Or CantAnios > AnioHasta Then
    Exit Sub
End If

MesDesde = MesHasta - CantMeses + 1
AnioDesde = AnioHasta - CantAnios


If MesDesde <= 0 Then
    MesDesde = MesHasta + 12 - CantMeses + 1
    AnioDesde = AnioDesde - 1
End If

If MesDesde > 12 Then
    MesDesde = MesDesde - 12
    AnioDesde = AnioDesde - 1
End If


' Modificado para que tome el promedio para los jornales
If ConFases Then
    'FGZ - 16/04/2004
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        'rs_Fases.MoveFirst
        rs_Fases.MoveLast
        If rs_Fases!altfec > CDate("01/" & MesDesde & "/" & AnioDesde) Then
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
        End If
    End If
Else
    'FGZ - 15/12/2004
    ' Se le saco esto porque traia problemas
'    If buliq_empleado!empfaltagr > CDate("01/" & MesDesde & "/" & AnioDesde) Then
'        MesDesde = Month(buliq_empleado!empfaltagr)
'        AnioDesde = Year(buliq_empleado!empfaltagr)
'    End If
End If

If TipoAcu = 0 Then 'Todos los tipos
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
Else
    If AnioDesde = AnioHasta Then
        If MesDesde = MesHasta Then
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames =" & MesHasta
        Else
            StrSql = "SELECT * FROM impmesarg " & _
                     " WHERE ternro = " & buliq_empleado!ternro & _
                     " AND impmesarg.tconnro =" & TipoAcu & _
                     " AND acunro =" & Acu & _
                     " AND " & AnioDesde & " = imaanio " & _
                     " AND imames >= " & MesDesde & " AND imames <=" & MesHasta
        End If
    Else
        StrSql = "SELECT * FROM impmesarg " & _
                 " WHERE ternro = " & buliq_empleado!ternro & _
                 " AND impmesarg.tconnro =" & TipoAcu & _
                 " AND acunro =" & Acu & _
                 " AND ((" & AnioDesde & " = imaanio AND imames >= " & MesDesde & ") OR " & _
                 " (imaanio > " & AnioDesde & " AND imaanio < " & AnioHasta & ") OR " & _
                 " (imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
    End If
End If
StrSql = StrSql & " ORDER BY imaanio, imames"
OpenRecordset StrSql, rs_ImpMes


'If TipoAcu = 0 Then 'Todos los tipos
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'Else
'        StrSql = "SELECT * FROM impmesarg " & _
'                 " WHERE ternro = " & buliq_empleado!ternro & _
'                 " AND impmesarg.tconnro =" & TipoAcu & _
'                 " AND acunro =" & Acu & _
'                 " AND " & AnioDesde & " <= imaanio AND imaanio <= " & AnioHasta & _
'                 " AND (( imames >=" & MesDesde & " AND " & AnioDesde & " <= imaanio) OR ( imames <=" & MesHasta & " AND imaanio = " & AnioHasta & "))"
'End If
'StrSql = StrSql & " ORDER BY imaanio, imames"
'OpenRecordset StrSql, rs_ImpMes

cantProm = 0
Do While Not rs_ImpMes.EOF
    If Not EsNulo(rs_ImpMes!imamonto) Then
        If rs_ImpMes!imamonto <> 0 Then ' si el monto es cero ==> no lo tengo en cuenta
            Valor = Valor + rs_ImpMes!imamonto
            Cantidad = Cantidad + IIf(Not EsNulo(rs_ImpMes!imacant), rs_ImpMes!imacant, 0)
            cantProm = cantProm + 1
        End If
    End If
    rs_ImpMes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    If TipoAcu = 0 Then 'Todos los tipos
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu
    Else
        StrSql = "SELECT * FROM impproarg WHERE cliqnro = " & buliq_cabliq!cliqnro & _
                 " AND acunro =" & Acu & " AND tconnro = " & TipoAcu
    End If
    OpenRecordset StrSql, rs_ImpPro
    If Not rs_ImpPro.EOF Then
        If Not EsNulo(rs_ImpPro!ipamonto) Then
            If rs_ImpPro!ipamonto <> 0 Then ' si el monto es cero ==> no lo tengo en cuenta
                Valor = Valor + rs_ImpPro!ipamonto
                Cantidad = Cantidad + IIf(Not EsNulo(rs_ImpPro!ipacant), rs_ImpPro!ipacant, 0)
                cantProm = cantProm + 1
            End If
        End If
    End If
End If

If cantProm <> 0 Then
    Valor = Valor / cantProm
    Cantidad = Cantidad / cantProm
Else
    Valor = 0
    Cantidad = 0
End If


' Cierro todo y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
If rs_ImpMes.State = adStateOpen Then rs_ImpMes.Close
If rs_ImpPro.State = adStateOpen Then rs_ImpPro.Close
Set rs_Fases = Nothing
Set rs_ImpMes = Nothing
Set rs_ImpPro = Nothing

End Sub


Public Sub bus_Campo0()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca un campo en la BD. gcampo0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroCampo As Long    'buscampo.bcnro

'Dim Param_cur As New ADODB.Recordset
Dim rs_campo As New ADODB.Recordset

    Bien = False
    Valor = 0

    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroCampo = Arr_Programa(NroProg).Auxint1
    Else
        Exit Sub
    End If

    StrSql = "SELECT * FROM buscampo WHERE bcnro = " & NroCampo
    OpenRecordset StrSql, rs_campo
    
    If Not rs_campo.EOF Then
        Bien = True
    End If


' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_campo.State = adStateOpen Then rs_campo.Close
Set rs_campo = Nothing
'Set Param_cur = Nothing

End Sub


Public Sub bus_Remun0()
' ---------------------------------------------------------------------------------------------
' Descripcion: Sumatoria de las Remuneraciones de las Ganancias, a partir de
'              Fecha de Pago � Mes/A�o del periodo. gremun0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim FecPago As Boolean  ' True  - Mes/A�o del Periodo
                        ' False - Fecha de Pago

'Dim Param_cur As New ADODB.Recordset
Dim rs_FichaRem As New ADODB.Recordset
    
'Definiciones Auxiliares
Dim Anio As Integer
Dim Rem_Per As Single
Dim ded_per As Single
Dim FechaPago As Date
Dim FechaAux As Date
Dim AuxDia As Integer
Dim AuxMes As Integer
Dim AuxAnio As Integer

    Bien = False
    Valor = 0

    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        FecPago = CBool(Arr_Programa(NroProg).Auxint1)
    Else
        Exit Sub
    End If

If FecPago Then
    If Not EsNulo(buliq_proceso!profecpago) Then
        FechaPago = buliq_proceso!profecpago
    Else
        FechaPago = Date ' La fecha de hoy
    End If
    
    AuxDia = 1
    If Month(FechaPago) = 12 Then
        AuxMes = 1
        AuxAnio = Year(FechaPago) + 1
    Else
        AuxMes = Month(FechaPago) + 1
        AuxAnio = Year(FechaPago) - 1
    End If
    
    FechaAux = CDate(AuxDia & "/" & AuxMes & "/" & AuxAnio)
Else
    AuxDia = 1
    If buliq_periodo!pliqmes = 12 Then
        AuxMes = 1
        AuxAnio = buliq_periodo!pliqanio + 1
    Else
        AuxMes = buliq_periodo!pliqmes + 1
        AuxAnio = buliq_periodo!pliqanio - 1
    End If
    
    FechaAux = CDate(AuxDia & "/" & AuxMes & "/" & AuxAnio)
    
    Anio = buliq_periodo!pliqanio
End If

' Calculo de Remuneaciones del a�o
If FecPago Then
    StrSql = "SELECT * FROM ficharem WHERE empleado = " & buliq_cabliq!Empleado & _
             " fecha <= " & FechaAux
Else
    StrSql = "SELECT * FROM ficharem WHERE empleado = " & buliq_cabliq!Empleado & _
             " fecha <= " & FechaAux & _
             " AND fecha >= " & CDate("01/01/" & Anio) & _
             " AND fecha <= " & CDate("31/12/" & Anio)
End If
OpenRecordset StrSql, rs_FichaRem

Rem_Per = 0
Do While Not rs_FichaRem.EOF
    Rem_Per = Rem_Per + rs_FichaRem!sujapor + rs_FichaRem!nsujapor + rs_FichaRem!nremgcias
        
    rs_FichaRem.MoveNext
Loop

Bien = True
Valor = Rem_Per

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_FichaRem.State = adStateOpen Then rs_FichaRem.Close
'Set Param_cur = Nothing
Set rs_FichaRem = Nothing

End Sub


Public Sub bus_Concep1()
' ---------------------------------------------------------------------------------------------
' Descripcion: Genera la novedad a nivel Global, Grupal o Individual. gconcep1.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.: FGZ - 28/09/2004
' Descripcion: vigencias
' ---------------------------------------------------------------------------------------------

Dim CodConc As Long       ' Nro de Concepto
Dim NroTpar As Long       ' Nro de parametro
Dim General As Boolean    ' Global
Dim Grupal As Boolean     ' Grupal
Dim Individual As Boolean ' Individual

Dim Firmado As Boolean
Dim Vigencia_Activa As Boolean
Dim Encontro As Boolean

Dim rs_NovEmp As New ADODB.Recordset
Dim rs_NovEstr As New ADODB.Recordset
Dim rs_NovGral As New ADODB.Recordset
Dim rs_firmas As New ADODB.Recordset
Dim rs_His_Estructura As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    General = False
    Grupal = False
    Individual = False
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        CodConc = Arr_Programa(NroProg).Auxint1
        NroTpar = Arr_Programa(NroProg).Auxint2
        General = CBool(Arr_Programa(NroProg).Auxlog3)
        Grupal = CBool(Arr_Programa(NroProg).Auxlog2)
        Individual = CBool(Arr_Programa(NroProg).Auxlog1)
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busca la novedad de otro concepto " & IIf(General, "Global ", "") & IIf(Grupal, "Estructura ", "") & IIf(Individual, "Individual ", "")
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la busqueda la Busqueda " & NroProg
        End If
        Exit Sub
    End If

    If Individual Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Individual, por empleado "
        End If
    
        StrSql = "SELECT * FROM novemp WHERE " & _
                 " concnro = " & CodConc & _
                 " AND tpanro = " & NroTpar & _
                 " AND empleado = " & buliq_cabliq!Empleado & _
                 " AND ((nevigencia = -1 " & _
                 " AND nedesde < " & ConvFecha(fecha_fin) & _
                 " AND (nehasta >= " & ConvFecha(fecha_inicio) & _
                 " OR nehasta is null )) " & _
                 " OR nevigencia = 0)" & _
                 " ORDER BY nevigencia, nedesde, nehasta "
        OpenRecordset StrSql, rs_NovEmp
        
        Valor = 0
        Do While Not rs_NovEmp.EOF
            If FirmaActiva5 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                         " AND cysfircodext = '" & rs_NovEmp!nenro & "' and cystipnro = 5"
                OpenRecordset StrSql, rs_firmas
                If rs_firmas.EOF Then
                    Firmado = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                    End If
                Else
                    Firmado = True
                End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If

            If Firmado Then
                If CBool(rs_NovEmp!nevigencia) Then
                    Vigencia_Activa = True
                    If Not EsNulo(rs_NovEmp!nehasta) Then
                        If (rs_NovEmp!nehasta < fecha_inicio) Or (fecha_fin < rs_NovEmp!nedesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " INACTIVA con valor " & rs_NovEmp!nevalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " ACTIVA con valor " & rs_NovEmp!nevalor
                            End If
                        End If
                    Else
                        If (fecha_fin < rs_NovEmp!nedesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEmp!nevalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEmp!nevalor
                            End If
                        End If
                    End If
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "Novedad sin vigencia con valor " & rs_NovEmp!nevalor
                    End If
                End If
                
                If Vigencia_Activa Or Not CBool(rs_NovEmp!nevigencia) Then
                    Valor = Valor + rs_NovEmp!nevalor
                End If
                If Not EsNulo(rs_NovEmp!neretro) Then
                    Retro = rs_NovEmp!neretro
                End If
                
                Encontro = True
            End If 'If Firmado Then
        
            rs_NovEmp.MoveNext
        Loop
        If Not Encontro Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad Individual "
            End If
        End If
        Bien = True
        Exit Sub
    End If
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
    
    If Grupal Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Por estructura "
        End If
    
        StrSql = "SELECT * FROM novestr WHERE " & _
                 " concnro = " & CodConc & _
                 " AND tpanro = " & NroTpar & _
                 " AND ((ntevigencia = -1 " & _
                 " AND ntedesde < " & ConvFecha(fecha_fin) & _
                 " AND (ntehasta >= " & ConvFecha(fecha_inicio) & " " & _
                 " OR ntehasta is null)) " & _
                 " OR ntevigencia = 0) " & _
                 " ORDER BY ntevigencia, ntedesde, ntehasta "
        OpenRecordset StrSql, rs_NovEstr
        
        Encontro = False
        If rs_NovEstr.EOF Then
            Firmado = False
        End If
        Valor = 0
        Do While Not rs_NovEstr.EOF 'And Not Encontro
            If FirmaActiva15 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovEstr!estrnro & "' and cystipnro = 15"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                        End If
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
            
            If Firmado Then
                If CBool(rs_NovEstr!ntevigencia) Then
                    Vigencia_Activa = True
                    If Not EsNulo(rs_NovEstr!ntehasta) Then
                        If (rs_NovEstr!ntehasta < fecha_inicio) Or (fecha_fin < rs_NovEstr!ntedesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta " & rs_NovEstr!ntehasta & " INACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta " & rs_NovEstr!ntehasta & " ACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        End If
                    Else
                        If (fecha_fin < rs_NovEstr!ntedesde) Then
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        End If
                    End If
                End If
                
                If Vigencia_Activa Or Not CBool(rs_NovEstr!ntevigencia) Then
                    Encontro = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "busco que el empleado tenga la estructura " & rs_NovEstr!estrnro & " activa"
                    End If
                    
                    'busco que el empleado tenga esa estructura activa
                    StrSql = " SELECT tenro, estrnro FROM his_estructura " & _
                             " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                             " estrnro =" & rs_NovEstr!estrnro & _
                             " AND (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND " & _
                             " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_His_Estructura
                    If Not rs_His_Estructura.EOF Then
                        Valor = Valor + rs_NovEstr!ntevalor
                        If Not EsNulo(rs_NovEstr!nteretro) Then
                            Retro = rs_NovEstr!nteretro
                        End If
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Estructura " & rs_NovEstr!estrnro & " activa"
                        End If
                        
                        Encontro = True
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Estructura " & rs_NovEstr!estrnro & " No activa"
                        End If
                    End If
                End If
            End If 'firmado
            
            rs_NovEstr.MoveNext
        Loop
        If Not Encontro Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad por Estructura"
            End If
        End If
        Bien = True
        Exit Sub
    End If
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

    If General Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Global "
        End If
    
        If objCache_NovGlobales.EsSimboloDefinido(CStr(CodConc & "-" & NroTpar)) Then
            Valor = objCache_NovGlobales.Valor(CStr(CodConc & "-" & NroTpar))
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Novedad en Cache con valor " & Valor
            End If
            
            Encontro = True
        Else
            StrSql = "SELECT * FROM novgral WHERE " & _
                     " concnro = " & CodConc & _
                     " AND tpanro = " & NroTpar & _
                     " AND ((ngravigencia = -1 " & _
                     " AND ngradesde < " & ConvFecha(fecha_fin) & " " & _
                     " AND (ngrahasta >= " & ConvFecha(fecha_inicio) & " " & _
                     " OR ngrahasta is null)) " & _
                     " OR ngravigencia = 0) " & _
                     " ORDER BY ngravigencia, ngradesde, ngrahasta "
            OpenRecordset StrSql, rs_NovGral
               
            Do While Not rs_NovGral.EOF
                If FirmaActiva19 Then
                    '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                        StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                                 " AND cysfircodext = '" & rs_NovGral!ngranro & "' and cystipnro = 19"
                        OpenRecordset StrSql, rs_firmas
                        If rs_firmas.EOF Then
                            Firmado = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                            End If
                        Else
                            Firmado = True
                        End If
                    If rs_firmas.State = adStateOpen Then rs_firmas.Close
                Else
                    Firmado = True
                End If
            
                If Firmado Then
                    If CBool(rs_NovGral!ngravigencia) Then
                        Vigencia_Activa = True
                        If Not EsNulo(rs_NovGral!ngrahasta) Then
                            If (rs_NovGral!ngrahasta < fecha_inicio) Or (fecha_fin < rs_NovGral!ngradesde) Then
                                 Vigencia_Activa = False
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " INACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            Else
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " ACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            End If
                        Else
                            If (fecha_fin < rs_NovGral!ngradesde) Then
                                 Vigencia_Activa = False
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado INACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            Else
                                If CBool(USA_DEBUG) Then
                                    Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado ACTIVA con valor " & rs_NovGral!ngravalor
                                End If
                            End If
                        End If
                    End If
                    
                    If Vigencia_Activa Or Not CBool(rs_NovGral!ngravigencia) Then
                        Valor = Valor + rs_NovGral!ngravalor
                        
                        If Not EsNulo(rs_NovGral!ngraretro) Then
                            Retro = rs_NovGral!ngraretro
                        End If
                        
                        Encontro = True
                    End If
                End If 'If Firmado Then
                
                rs_NovGral.MoveNext
            Loop
        End If
        If Not Encontro Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad global"
            End If
        Else
            'inserto la novedad en el cache
            Call objCache_NovGlobales.Insertar_Simbolo(CStr(CodConc & "-" & NroTpar & "0"), Valor)
        End If
        Bien = True
        Exit Sub
    End If

' cierro todo y libero
    If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
    If rs_NovEstr.State = adStateOpen Then rs_NovEstr.Close
    If rs_NovGral.State = adStateOpen Then rs_NovGral.Close
    If rs_firmas.State = adStateOpen Then rs_firmas.Close
    If rs_His_Estructura.State = adStateOpen Then rs_His_Estructura.Close
        
    Set rs_NovEmp = Nothing
    Set rs_NovEstr = Nothing
    Set rs_NovGral = Nothing
    Set rs_firmas = Nothing
    Set rs_His_Estructura = Nothing
    
End Sub


Public Sub bus_Prestamos()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de los prestamos de cualquier tipo de prestamos.
'              Mensuales, 1era o 2da quincena.
'       gprest00.p
' Autor      : FGZ
' Fecha      : 04/06/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Final As Boolean    'Liquidacion Final
Dim Cancela As Boolean  'Si cancela las cuaotas y prestamos o no
Dim Nrotp As Long       'Tipo de Prestamos
Dim CodMone As Integer  ' Moneda.monnro
Dim Opcion As Integer   ' 1 - Mensual
                        ' 2 - 1er Quincena
                        ' 3 - 2da Quincena

'Dim Param_cur As New ADODB.Recordset
Dim rs_Prestamo As New ADODB.Recordset
Dim rs_Cuota As New ADODB.Recordset
Dim rs_Aux_Cuota As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        If EsNulo(Arr_Programa(NroProg).Auxint1) Then
            Nrotp = -1 ' Todos
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Todos los Acumuladores "
            End If
        Else
            If Arr_Programa(NroProg).Auxint1 = 0 Then
                Nrotp = -1 ' Todos
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Todos los Tipos de Prestamos "
                End If
            Else
                Nrotp = Arr_Programa(NroProg).Auxint1
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Acumulador " & Nrotp
                End If
            End If
        End If
        
        If EsNulo(Arr_Programa(NroProg).Auxint2) Then
            CodMone = -1 ' Todas
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Para Todas las Monedas "
            End If
        Else
            If Arr_Programa(NroProg).Auxint2 = 0 Then
                CodMone = -1 ' Todas
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Para Todas las Monedas "
                End If
            Else
                CodMone = Arr_Programa(NroProg).Auxint2
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Para la Moneda " & CodMone
                End If
            End If
        End If
        
        Opcion = Arr_Programa(NroProg).Auxint3
        
        Final = CBool(Arr_Programa(NroProg).Auxlog1)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros cargados "
        End If
        
        Cancela = IIf(Not EsNulo(Arr_Programa(NroProg).Auxlog2), CBool(Arr_Programa(NroProg).Auxlog2), True)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la busqueda " & NroProg
        End If
        Exit Sub
    End If


If Final Then 'se trata de una liq. final se descuentan todas las cuotas
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "se trata de una liq. final se descuentan todas las cuotas "
    End If

    Select Case Opcion
    Case 1: 'Prestamos Mensuales
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los prestamos mensuales "
        End If
        StrSql = "SELECT * FROM prestamo "
        StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
        StrSql = StrSql & " WHERE prestamo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  prestamo.estnro = 3 "
        StrSql = StrSql & " AND prestamo.quincenal = 0 "
        If CodMone <> -1 Then
            StrSql = StrSql & " AND prestamo.monnro =" & CodMone
        End If
        If Nrotp <> -1 Then
            StrSql = StrSql & " AND pre_linea.tpnro =" & Nrotp
        End If
        OpenRecordset StrSql, rs_Prestamo
        
        If CBool(USA_DEBUG) Then
            If rs_Prestamo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron prestamos "
            End If
        End If
        Do While Not rs_Prestamo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el prestamo " & rs_Prestamo!prenro
            End If
            StrSql = "SELECT * FROM pre_cuota "
            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!cuototal
                
                If Cancela Then
                    StrSql = "UPDATE pre_cuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", cuocancela = -1 "
                    StrSql = StrSql & " WHERE pre_cuota.cuonro = " & rs_Cuota!cuonro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                Bien = True
                rs_Cuota.MoveNext
            Loop
            
'            StrSql = "SELECT * FROM pre_cuota "
'            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
'            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
'            OpenRecordset StrSql, rs_Aux_Cuota
'            If rs_Aux_Cuota.EOF Then
            If Cancela Then
                StrSql = "UPDATE prestamo SET estnro = 6"
                StrSql = StrSql & " WHERE prenro = " & rs_Prestamo!prenro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
'            End If
            
            rs_Prestamo.MoveNext
        Loop
        
    Case 2, 3: 'Prestamos de la Primera y Segunda Quincena
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los prestamos de la primera y segunda quincena "
        End If
    
        StrSql = "SELECT * FROM prestamo "
        StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
        StrSql = StrSql & " WHERE prestamo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  prestamo.estnro = 3 "
        StrSql = StrSql & " AND prestamo.quincenal = -1 "
        If CodMone <> -1 Then
            StrSql = StrSql & " AND prestamo.monnro =" & CodMone
        End If
        If Nrotp <> -1 Then
            StrSql = StrSql & " AND pre_linea.tpnro =" & Nrotp
        End If
        OpenRecordset StrSql, rs_Prestamo
        
        If CBool(USA_DEBUG) Then
            If rs_Prestamo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron prestamos "
            End If
        End If
        
        Do While Not rs_Prestamo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el prestamo " & rs_Prestamo!prenro
            End If
            StrSql = "SELECT * FROM pre_cuota "
            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!cuototal
                
                If Cancela Then
                    StrSql = "UPDATE pre_cuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", cuocancela = -1 "
                    StrSql = StrSql & " WHERE pre_cuota.cuonro = " & rs_Cuota!cuonro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
            
'            StrSql = "SELECT * FROM pre_cuota "
'            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
'            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
'            OpenRecordset StrSql, rs_Aux_Cuota
'            If rs_Aux_Cuota.EOF Then
            If Cancela Then
                StrSql = "UPDATE prestamo SET estnro = 6"
                StrSql = StrSql & " WHERE prenro = " & rs_Prestamo!prenro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
'            End If
            
            rs_Prestamo.MoveNext
        Loop
    End Select
Else 'liquidacion mensual o quincenal
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "se trata de una liquidacion mensual o quincenal "
    End If

    Select Case Opcion
    Case 1: 'Prestamos Mensuales
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los prestamos mensuales "
        End If
        StrSql = "SELECT * FROM prestamo "
        StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
        StrSql = StrSql & " WHERE prestamo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  prestamo.estnro = 3 "
        StrSql = StrSql & " AND prestamo.quincenal = 0 "
        If CodMone <> -1 Then
            StrSql = StrSql & " AND prestamo.monnro =" & CodMone
        End If
        If Nrotp <> -1 Then
            StrSql = StrSql & " AND pre_linea.tpnro =" & Nrotp
        End If
        OpenRecordset StrSql, rs_Prestamo
        
        If CBool(USA_DEBUG) Then
            If rs_Prestamo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron prestamos "
            End If
        End If
        
        Do While Not rs_Prestamo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el prestamo " & rs_Prestamo!prenro
            End If
        
            StrSql = "SELECT * FROM pre_cuota "
            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
            StrSql = StrSql & " AND pre_cuota.cuoano = " & buliq_periodo!pliqanio
            StrSql = StrSql & " AND pre_cuota.cuomes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!cuototal
                
                If Cancela Then
                    StrSql = "UPDATE pre_cuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", cuocancela = -1 "
                    StrSql = StrSql & " WHERE pre_cuota.cuonro = " & rs_Cuota!cuonro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
            
'            StrSql = "SELECT * FROM pre_cuota "
'            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
'            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
'            OpenRecordset StrSql, rs_Aux_Cuota
'            If rs_Aux_Cuota.EOF Then
'                StrSql = "UPDATE prestamo SET estnro = 6"
'                StrSql = StrSql & " WHERE prenro = " & rs_Prestamo!prenro
'                StrSql = StrSql & " AND pre_cuota.cuoano = " & buliq_periodo!pliqanio
'                StrSql = StrSql & " AND pre_cuota.cuomes = " & buliq_periodo!pliqmes
'                objConn.Execute StrSql, , adExecuteNoRecords
'            End If
            
            rs_Prestamo.MoveNext
        Loop
        
    Case 2: 'Prestamos de la Primera quincena
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los prestamos de la primera quincena "
        End If
    
        StrSql = "SELECT * FROM prestamo "
        StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
        StrSql = StrSql & " WHERE prestamo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  prestamo.estnro = 3 "
        StrSql = StrSql & " AND prestamo.quincenal = -1 "
        If CodMone <> -1 Then
            StrSql = StrSql & " AND prestamo.monnro =" & CodMone
        End If
        If Nrotp <> -1 Then
            StrSql = StrSql & " AND pre_linea.tpnro =" & Nrotp
        End If
        OpenRecordset StrSql, rs_Prestamo
        
        If CBool(USA_DEBUG) Then
            If rs_Prestamo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron prestamos "
            End If
        End If
        Do While Not rs_Prestamo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el prestamo " & rs_Prestamo!prenro
            End If
        
            StrSql = "SELECT * FROM pre_cuota "
            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
            StrSql = StrSql & " AND pre_cuota.cuoquin = 1 "
            StrSql = StrSql & " AND pre_cuota.cuoano = " & buliq_periodo!pliqanio
            StrSql = StrSql & " AND pre_cuota.cuomes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!cuototal
                
                If Cancela Then
                    StrSql = "UPDATE pre_cuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", cuocancela = -1 "
                    StrSql = StrSql & " WHERE pre_cuota.cuonro = " & rs_Cuota!cuonro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
            
'            StrSql = "SELECT * FROM pre_cuota "
'            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
'            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
'            OpenRecordset StrSql, rs_Aux_Cuota
'            If rs_Aux_Cuota.EOF Then
'                StrSql = "UPDATE prestamo SET estnro = 6"
'                StrSql = StrSql & " WHERE prenro = " & rs_Prestamo!prenro
'                objConn.Execute StrSql, , adExecuteNoRecords
'            End If
            
            rs_Prestamo.MoveNext
        Loop
    Case 3: 'Prestamos de la Segunda quincena
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los prestamos de la segunda quincena "
        End If
    
        StrSql = "SELECT * FROM prestamo "
        StrSql = StrSql & " INNER JOIN pre_linea ON pre_linea.lnprenro = prestamo.lnprenro "
        StrSql = StrSql & " WHERE prestamo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  prestamo.estnro = 3 "
        StrSql = StrSql & " AND prestamo.quincenal = -1 "
        If CodMone <> -1 Then
            StrSql = StrSql & " AND prestamo.monnro =" & CodMone
        End If
        If Nrotp <> -1 Then
            StrSql = StrSql & " AND pre_linea.tpnro =" & Nrotp
        End If
        OpenRecordset StrSql, rs_Prestamo
        
        If CBool(USA_DEBUG) Then
            If rs_Prestamo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron prestamos "
            End If
        End If
        
        Do While Not rs_Prestamo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el prestamo " & rs_Prestamo!prenro
            End If
        
            StrSql = "SELECT * FROM pre_cuota "
            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
            StrSql = StrSql & " AND pre_cuota.cuoquin = 2 "
            StrSql = StrSql & " AND pre_cuota.cuoano = " & buliq_periodo!pliqanio
            StrSql = StrSql & " AND pre_cuota.cuomes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!cuototal
                
                If Cancela Then
                    StrSql = "UPDATE pre_cuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", cuocancela = -1 "
                    StrSql = StrSql & " WHERE pre_cuota.cuonro = " & rs_Cuota!cuonro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
            
'            StrSql = "SELECT * FROM pre_cuota "
'            StrSql = StrSql & " WHERE pre_cuota.prenro =" & rs_Prestamo!prenro
'            StrSql = StrSql & " AND pre_cuota.cuocancela = 0 "
'            OpenRecordset StrSql, rs_Aux_Cuota
'            If rs_Aux_Cuota.EOF Then
'                StrSql = "UPDATE prestamo SET estnro = 6"
'                StrSql = StrSql & " WHERE prenro = " & rs_Prestamo!prenro
'                objConn.Execute StrSql, , adExecuteNoRecords
'            End If
            
            rs_Prestamo.MoveNext
        Loop
    End Select
End If

' cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_Cuota.State = adStateOpen Then rs_Cuota.Close
If rs_Prestamo.State = adStateOpen Then rs_Prestamo.Close
If rs_Aux_Cuota.State = adStateOpen Then rs_Aux_Cuota.Close
        
'Set Param_cur = Nothing
Set rs_Prestamo = Nothing
Set rs_Cuota = Nothing
Set rs_Aux_Cuota = Nothing

End Sub


Public Sub bus_Anti0(ByRef antdia As Integer, ByRef antmes As Integer, ByRef antanio As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca la Antiguedad de un Empleado, dependiendo de si es con tope o no,
'               o si es a una fecha o no. ganti0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Tope As Integer         ' 1 - con tope
                            ' 2 - sin tope
                            
Dim Fecha As Date           ' Fecha Hasta
Dim Tipo_Fase As Integer    ' 1 - Sueldo
                            ' 2 - Indemnisacion
                            ' 3 - Vacaciones
                            ' 4 - Real
Dim Hasta As Integer        ' 1 - A fin de Mes
                            ' 2 - A la fecha
                            ' 3 - A la fecha de baja
                            ' 4 - A Fecha de baja prevista
                            ' 5 - A Fin de a�o
                            
' estao es nuevo (10/11/2003)
Dim Cant_Tope As Integer    ' la cantidad de tope
Dim Resultado As Integer    ' 1 - en dias
                            ' 2 - en meses
                            ' 3 - en a�os
'Dim antdia As Integer
'Dim antmes As Integer
'Dim antanio As Integer
Dim q As Integer

Dim FechaAux As Date

'Dim Param_cur As New ADODB.Recordset
Dim rs_Fases  As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Tope = CInt(Arr_Programa(NroProg).Auxint1)
        Hasta = CInt(Arr_Programa(NroProg).Auxint2)
        Cant_Tope = CInt(Arr_Programa(NroProg).Auxint3)
        
        If Not EsNulo(Arr_Programa(NroProg).Auxchar1) Then
            Fecha = CDate(Arr_Programa(NroProg).Auxchar1)
        End If
        
        Tipo_Fase = CInt(Arr_Programa(NroProg).Auxint4)
        Resultado = CInt(Arr_Programa(NroProg).Auxint5)
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busca la Antiguedad de un Empleado " & IIf(Tope = 1, "con tope ", "sin tope ")
            Select Case Hasta
                Case 1:
                    StrLog = "a fin de mes "
                Case 2:
                    StrLog = "a la fecha " & Fecha
                Case 3:
                    StrLog = "a la fecha de baja de la fase "
                    Select Case Tipo_Fase
                        Case 1:
                            StrLog = StrLog & " Sueldo "
                        Case 2:
                            StrLog = StrLog & " Indemnizacion "
                        Case 3:
                            StrLog = StrLog & " Vacacion "
                        Case 4:
                        StrLog = StrLog & " real "
                    End Select
                Case 4:
                    StrLog = "a la fecha de baja prevista "
                Case 5:
                    StrLog = "a la finde a�o "
            End Select
            Flog.writeline Espacios(Tabulador * 4) & StrLog
            Select Case Resultado
                Case 1:
                    Flog.writeline Espacios(Tabulador * 4) & "el resultado en dias "
                Case 2:
                    Flog.writeline Espacios(Tabulador * 4) & "el resultado en meses "
                Case 3:
                    Flog.writeline Espacios(Tabulador * 4) & "el resultado en a�os "
            End Select
        End If
        
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If

    Select Case Tipo_Fase
    Case 1: 'Sueldo
            Select Case Hasta
            Case 1: 'Fin de mes
                If buliq_periodo!pliqmes = 12 Then
                    FechaAux = CDate("1/1/" & buliq_periodo!pliqanio + 1) - 1
                Else
                    FechaAux = CDate("01/" & buliq_periodo!pliqmes + 1 & "/" & buliq_periodo!pliqanio) - 1
                End If
                Call bus_Antiguedad("SUELDO", FechaAux, antdia, antmes, antanio, q)
            Case 2: 'A la Fecha
                Call bus_Antiguedad("SUELDO", Fecha, antdia, antmes, antanio, q)
            Case 3: 'A fecha de Baja
                StrSql = "SELECT * FROM fases WHERE estado = 0 AND empleado = " & buliq_empleado!ternro & _
                         " AND sueldo = -1 " & _
                         " AND not altfec is null " & _
                         " AND not bajfec is null " & _
                         " ORDER BY altfec "
                OpenRecordset StrSql, rs_Fases
                If Not rs_Fases.EOF Then
                    rs_Fases.MoveLast
                    FechaAux = rs_Fases!bajfec
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontro ninguna fase cerrada "
                    End If
                    Exit Sub
                End If
                Call bus_Antiguedad("SUELDO", FechaAux, antdia, antmes, antanio, q)
            Case 4: 'A fecha de baja prevista
                If Not EsNulo(buliq_empleado!empfbajaprev) Then
                    FechaAux = buliq_empleado!empfbajaprev
                    Call bus_Antiguedad("SUELDO", FechaAux, antdia, antmes, antanio, q)
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "Fecha de baja prevista nula "
                    End If
                    Exit Sub
                End If
            Case 5: 'A fin de a�o
                    FechaAux = CDate("31/12/" & buliq_periodo!pliqanio)
                    Call bus_Antiguedad("SUELDO", FechaAux, antdia, antmes, antanio, q)
            End Select
    Case 2: 'Indemnizaci�n
            Select Case Hasta
            Case 1: 'Fin de mes
                If buliq_periodo!pliqmes = 12 Then
                    FechaAux = CDate("1/1/" & buliq_periodo!pliqanio + 1) - 1
                Else
                    FechaAux = CDate("01/" & buliq_periodo!pliqmes + 1 & "/" & buliq_periodo!pliqanio) - 1
                End If
                    
                Call bus_Antiguedad("INDEMNIZACION", FechaAux, antdia, antmes, antanio, q)
            Case 2: 'A la Fecha
                Call bus_Antiguedad("INDEMNIZACION", Fecha, antdia, antmes, antanio, q)
            Case 3: 'A fecha de Baja
                StrSql = "SELECT * FROM fases WHERE estado = 0 AND empleado = " & buliq_empleado!ternro & _
                         " AND indemnizacion = -1 " & _
                         " AND not altfec is null " & _
                         " AND not bajfec is null " & _
                         " ORDER BY altfec "
                OpenRecordset StrSql, rs_Fases
                If Not rs_Fases.EOF Then
                    rs_Fases.MoveLast
                    FechaAux = rs_Fases!bajfec
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontro ninguna fase cerrada "
                    End If
                    Exit Sub
                End If
                Call bus_Antiguedad("INDEMNIZACION", FechaAux, antdia, antmes, antanio, q)
            Case 4: 'A fecha de baja prevista
                If Not EsNulo(buliq_empleado!empfbajaprev) Then
                    FechaAux = buliq_empleado!empfbajaprev
                    Call bus_Antiguedad("INDEMNIZACION", FechaAux, antdia, antmes, antanio, q)
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "Fecha de baja prevista nula "
                    End If
                    Exit Sub
                End If
            Case 5: 'A fin de a�o
                    FechaAux = CDate("31/12/" & buliq_periodo!pliqanio)
                    Call bus_Antiguedad("INDEMNIZACION", FechaAux, antdia, antmes, antanio, q)
            End Select
    Case 3: 'Vacaciones
            Select Case Hasta
            Case 1: 'Fin de mes
                If buliq_periodo!pliqmes = 12 Then
                    FechaAux = CDate("1/1/" & buliq_periodo!pliqanio + 1) - 1
                Else
                    FechaAux = CDate("01/" & buliq_periodo!pliqmes + 1 & "/" & buliq_periodo!pliqanio) - 1
                End If
                    
                Call bus_Antiguedad("VACACIONES", FechaAux, antdia, antmes, antanio, q)
            Case 2: 'A la Fecha
                Call bus_Antiguedad("VACACIONES", Fecha, antdia, antmes, antanio, q)
            Case 3: 'A fecha de Baja
                StrSql = "SELECT * FROM fases WHERE estado = 0 AND empleado = " & buliq_empleado!ternro & _
                         " AND vacaciones = -1 " & _
                         " AND not altfec is null " & _
                         " AND not bajfec is null " & _
                         " ORDER BY altfec "
                OpenRecordset StrSql, rs_Fases
                If Not rs_Fases.EOF Then
                    rs_Fases.MoveLast
                    FechaAux = rs_Fases!bajfec
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontro ninguna fase cerrada "
                    End If
                    Exit Sub
                End If
                Call bus_Antiguedad("VACACIONES", FechaAux, antdia, antmes, antanio, q)
            Case 4: 'A fecha de baja prevista
                If Not EsNulo(buliq_empleado!empfbajaprev) Then
                    FechaAux = buliq_empleado!empfbajaprev
                    Call bus_Antiguedad("VACACIONES", FechaAux, antdia, antmes, antanio, q)
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "Fecha de baja prevista nula "
                    End If
                    Exit Sub
                End If
            Case 5: 'A fin de a�o
                    FechaAux = CDate("31/12/" & buliq_periodo!pliqanio)
                    Call bus_Antiguedad("VACACIONES", FechaAux, antdia, antmes, antanio, q)
            End Select
    Case 4: 'Real
            Select Case Hasta
            Case 1: 'Fin de mes
                If buliq_periodo!pliqmes = 12 Then
                    FechaAux = CDate("1/1/" & buliq_periodo!pliqanio + 1) - 1
                Else
                    FechaAux = CDate("01/" & buliq_periodo!pliqmes + 1 & "/" & buliq_periodo!pliqanio) - 1
                End If
                    
                Call bus_Antiguedad("REAL", FechaAux, antdia, antmes, antanio, q)
            Case 2: 'A la Fecha
                Call bus_Antiguedad("REAL", Fecha, antdia, antmes, antanio, q)
            Case 3: 'A fecha de Baja
                StrSql = "SELECT * FROM fases WHERE estado = 0 AND empleado = " & buliq_empleado!ternro & _
                         " AND real = -1 " & _
                         " AND not altfec is null " & _
                         " AND not bajfec is null " & _
                         " ORDER BY altfec "
                OpenRecordset StrSql, rs_Fases
                If Not rs_Fases.EOF Then
                    rs_Fases.MoveLast
                    FechaAux = rs_Fases!bajfec
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontro ninguna fase cerrada "
                    End If
                    Exit Sub
                End If
                Call bus_Antiguedad("REAL", FechaAux, antdia, antmes, antanio, q)
            Case 4: 'A fecha de baja prevista
                If Not EsNulo(buliq_empleado!empfbajaprev) Then
                    FechaAux = buliq_empleado!empfbajaprev
                    Call bus_Antiguedad("REAL", FechaAux, antdia, antmes, antanio, q)
                Else
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "Fecha de baja prevista nula "
                    End If
                    Exit Sub
                End If
            Case 5: 'A fin de a�o
                    FechaAux = CDate("31/12/" & buliq_periodo!pliqanio)
                    Call bus_Antiguedad("REAL", FechaAux, antdia, antmes, antanio, q)
            End Select
    Case Else
    End Select
    
    ' FGZ
    ' 10/1/2003
    ' esto es nuevo
    Select Case Resultado
    Case 1: ' En dias
        Valor = antdia + antmes * 30 + antanio * 360
    Case 2: ' En meses
        Valor = antmes + antanio * 12
    Case 3: ' en a�os
        Valor = antanio
    End Select
         
    If Tope = 1 Then
        If Valor > Cant_Tope Then
            Valor = Cant_Tope
        End If
    End If
         
    Bien = True
    
' cierro todo y libero
    'If Param_cur.State = adStateOpen Then Param_cur.Close
    'Set Param_cur = Nothing
End Sub

Public Sub bus_Antiguedad(ByVal TipoAnt As String, ByVal FechaFin As Date, ByRef dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias h�biles(si es menor que un a�o) o en dias, meses y a�os en caso contrario.
'              antiguedad.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer

Dim NombreCampo As String

Dim rs_Fases As New ADODB.Recordset


NombreCampo = ""
DiasHabiles = 0

Select Case UCase(TipoAnt)
Case "SUELDO":
    NombreCampo = "sueldo"
Case "INDEMNIZACION":
    NombreCampo = "indemnizacion"
Case "VACACIONES":
    NombreCampo = "vacaciones"
Case "REAL":
    NombreCampo = "real"
Case Else
End Select

'StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
'         " AND " & NombreCampo & " = -1 "
'OpenRecordset StrSql, rs_Fases

' FGZ -27/01/2004
StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
         " AND " & NombreCampo & " = -1 " & _
         " AND not altfec is null " & _
         " AND not (bajfec is null AND estado = 0)" & _
         " AND altfec <= " & ConvFecha(fecha_fin)
OpenRecordset StrSql, rs_Fases

Do While Not rs_Fases.EOF
'    If (EsNulo(rs_Fases!altfec)) Or (EsNulo(rs_Fases!bajfec) And rs_Fases!estado = 0) Or (rs_Fases!altfec >= FechaFin) Then
'        GoTo siguiente
'    Else
        fecalta = rs_Fases!altfec
'    End If
    
    ' Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If rs_Fases!estado Then
        fecbaja = FechaFin ' solo es un alta, tomar el fecha-fin
    ElseIf rs_Fases!bajfec <= FechaFin Then
        fecbaja = rs_Fases!bajfec  ' se trata de un registro completo
    Else
        fecbaja = FechaFin ' hasta la fecha ingresada
    End If
    
'    Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
'    Dia = Dia + aux1
'    Mes = Mes + aux2 + Int(Dia / 30)
'    Anio = Anio + aux3 + Int(Mes / 12)
'    Dia = Dia Mod 30
'    Mes = Mes Mod 12
        
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "fase de " & fecalta & " a " & fecbaja
    End If
        
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    'Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
    If rs_Fases.RecordCount = 1 Then
        dia = aux1
        Mes = aux2
        Anio = aux3
    Else
        dia = dia + aux1
        Mes = Mes + aux2 + Int(dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        dia = dia Mod 30
        Mes = Mes Mod 12
    End If
        
    If Anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
siguiente:
    rs_Fases.MoveNext
Loop

If Anio <> 0 Then
    DiasHabiles = 0
End If


' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub

Public Sub bus_Antiguedad_A_FechaAlta(ByRef dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad al dia de hoy de un empleado en :
'              dias h�biles(si es menor que un a�o) o en dias, meses y a�os en caso contrario.
' Autor      : FGZ
' Fecha      : 21/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer
Dim DiasHabiles As Integer

DiasHabiles = 0

    fecalta = CDate(buliq_empleado!empfaltagr)
    fecbaja = CDate("31/12/" & buliq_periodo!pliqanio)
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Desde " & fecalta & " a " & fecbaja
    End If
        
    Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
    dia = dia + aux1
    Mes = Mes + aux2 + Int(dia / 30)
    Anio = Anio + aux3 + Int(Mes / 12)
    dia = dia Mod 30
    Mes = Mes Mod 12
        
    If Anio = 0 Then
        Call DiasTrab(fecalta, fecbaja, aux1)
        DiasHabiles = DiasHabiles + aux1
    End If
    
If Anio <> 0 Then
    DiasHabiles = 0
End If

'resultado en meses
Valor = Mes + Anio * 12

'    Select Case Resultado
'    Case 1: ' En dias
'        Valor = antdia + antmes * 30 + antanio * 360
'    Case 2: ' En meses
'        Valor = antmes + antanio * 12
'    Case 3: ' en a�os
'        Valor = antanio
'    End Select
'
'    If Tope = 1 Then
'        If Valor > Cant_Tope Then
'            Valor = Cant_Tope
'        End If
'    End If
Bien = True
End Sub


'Public Sub bus_Anti1()
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Busca la Antiguedad de un Empleado, dependiendo de si es con tope o no,
''               o si es a una fecha o no. ganti1.p
'' Autor      : FGZ
'' Fecha      :
'' Ultima Mod.:
'' Descripcion:
'' ---------------------------------------------------------------------------------------------
'
'Dim Tope As Integer         ' 1 - Con Tope
'                            ' 2 - Sin Tope
'Dim fecha As Date           ' Fecha Hasta
'Dim Hasta As Integer        ' 1 - A fin de Mes
'                            ' 2 - A la fecha
'
'' estao es nuevo (10/11/2003)
'Dim Cant_Tope As Integer    ' la cantidad de tope
'Dim Resultado As Integer    ' 1 - en dias
'                            ' 2 - en meses
'                            ' 3 - en a�os
'
'Dim antdia As Integer
'Dim antmes As Integer
'Dim antanio As Integer
'Dim q As Integer
'
'Dim FechaAux As Date
'
''Dim Param_cur As New ADODB.Recordset
'
'    Bien = False
'    valor = 0
'
'    ' Obtener los parametros de la Busqueda
'    strsql = "SELECT auxint1, auxlog1, auxlog2, auxlog3 FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        Tope = CInt(Arr_Programa(nroprog).auxint1)
'        Hasta = Arr_Programa(nroprog).auxint2
'        Cant_Tope = CInt(Arr_Programa(nroprog).auxint3)
'        If Not EsNulo(Arr_Programa(nroprog).auxchar1) Then
'            fecha = CDate(Arr_Programa(nroprog).auxchar1)
'        End If
'    Else
'        Exit Sub
'    End If
'
'    'Vacaciones
'    If Not Hasta Then
'        If buliq_periodo!pliqmes = 12 Then
'            FechaAux = CDate("1/1/" & buliq_periodo!pliqanio)
'        Else
'            FechaAux = CDate(buliq_periodo!pliqmes + 1 & "/1/" & buliq_periodo!pliqanio - 1)
'        End If
'
'        Call bus_Antiguedad("VACACIONES", FechaAux, antdia, antmes, antanio, q)
'    Else
'        Call bus_Antiguedad("VACACIONES", fecha, antdia, antmes, antanio, q)
'    End If
'
'    If antmes > 3 Then
'        If antanio > Tope Then
'            valor = Tope
'        Else
'            valor = antanio ' antiguedad en a�os
'        End If
'    Else
'        valor = antanio 'Antiguedad en a�os
'    End If
'    Bien = True
'
'
'' cierro todo y libero
'    'If Param_cur.State = adStateOpen Then Param_cur.Close
'    'Set Param_cur = Nothing
'End Sub
'
'
'Public Sub bus_Anti2()
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Busca la Antiguedad de un Empleado, dependiendo de si es con tope o no,
''               o si es a una fecha o no. ganti2.p
'' Autor      : FGZ
'' Fecha      :
'' Ultima Mod.:
'' Descripcion:
'' ---------------------------------------------------------------------------------------------
'
'Dim Tope As Boolean         ' si usa tope o no
'Dim fecha As Date           ' Fecha Hasta
'Dim Hasta As Integer        ' 1 - A fin de Mes
'                            ' 2 - A la fecha
'
'Dim antdia As Integer
'Dim antmes As Integer
'Dim antanio As Integer
'Dim q As Integer
'
'Dim FechaAux As Date
'
''Dim Param_cur As New ADODB.Recordset
'
'    Bien = False
'    valor = 0
'
'    ' Obtener los parametros de la Busqueda
'    strsql = "SELECT auxint1, auxlog1, auxlog2, auxlog3 FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        Tope = CBool(Arr_Programa(nroprog).auxint1)
'        Hasta = Arr_Programa(nroprog).auxint2
'
'        If Not EsNulo(Arr_Programa(nroprog).auxchar1) Then
'            fecha = CDate(Arr_Programa(nroprog).auxchar1)
'        End If
'    Else
'        Exit Sub
'    End If
'
'    'Indemnizacion
'    If Not Hasta Then
'        If buliq_periodo!pliqmes = 12 Then
'            FechaAux = CDate("1/1/" & buliq_periodo!pliqanio)
'        Else
'            FechaAux = CDate(buliq_periodo!pliqmes + 1 & "/1/" & buliq_periodo!pliqanio - 1)
'        End If
'
'        Call bus_Antiguedad("INDEMNIZACION", FechaAux, antdia, antmes, antanio, q)
'    Else
'        Call bus_Antiguedad("INDEMNIZACION", fecha, antdia, antmes, antanio, q)
'    End If
'
'    If antmes > 3 Then
'        If antanio > Tope Then
'            valor = Tope
'        Else
'            valor = antanio ' antiguedad en a�os
'        End If
'    Else
'        valor = antanio 'Antiguedad en a�os
'    End If
'    Bien = True
'
'
'' cierro todo y libero
'    'If Param_cur.State = adStateOpen Then Param_cur.Close
'    'Set Param_cur = Nothing
'End Sub

Public Sub DiasTrab(ByVal Desde As Date, ByVal Hasta As Date, ByRef DiasH As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(Hasta)
    
    aux = DateDiff("d", Desde, Hasta) + 1
    If aux < 7 Then
        DiasH = Minimo(aux, dxsem)
    Else
        If aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais.EOF Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(Hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estructura ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & buliq_empleado!ternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub

Public Sub DiasTrab_old(ByVal Desde As Date, ByVal Hasta As Date, ByRef DiasH As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la cantidad de dias trabajados de acuerdo al turno en que se trabaja y
'              de acuerdo a los dias que figuran como feriados en la tabla de feriados.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim d1 As Integer
Dim d2 As Integer
Dim aux As Integer
Dim aux2 As Integer
Dim dxsem As Integer

Dim rs_pais As New ADODB.Recordset
Dim rs_feriados As New ADODB.Recordset

    dxsem = 5
    
    d1 = Weekday(Desde)
    d2 = Weekday(Hasta)
    
    aux = DateDiff("d", Hasta, Desde) + 1
    If aux < 7 Then
        DiasH = Minimo(aux, dxsem)
    Else
        If aux = 7 Then
            DiasH = dxsem
        Else
            aux2 = 8 - d1 + d2
            If aux2 < 7 Then
                aux2 = Minimo(aux2, dxsem)
            Else
                If aux2 = 7 Then
                    aux2 = dxsem
                End If
            End If
            
            If aux2 >= 7 Then
                aux2 = Abs(aux2 - 7) + Int(aux2 / 7) * dxsem
            Else
                aux2 = aux2 + Int((aux2 - aux2) / 7) * dxsem
            End If
        End If
    End If
    
    aux = 0
    
    StrSql = "SELECT * FROM pais INNER JOIN tercero ON tercero.paisnro = pais.paisnro WHERE tercero.ternro = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_pais
    
    If Not rs_pais Then
        ' Resto los Feriados Nacionales
        StrSql = "SELECT * FROM feriado WHERE tipferinro = 2 " & _
                 " AND fericodext = " & rs_pais!paisnro & _
                 " AND ferifecha >= " & ConvFecha(Desde) & _
                 " AND ferifecha < " & ConvFecha(Hasta)
        OpenRecordset StrSql, rs_feriados
        
        Do While Not rs_feriados.EOF
            If Weekday(rs_feriados!ferifecha) > 1 Then
                DiasH = DiasH - 1
            End If
            
            ' Siguiente Feriado
            rs_feriados.MoveNext
        Loop
    End If


    ' Resto los feriados por Convenio
    StrSql = "SELECT * FROM empleado INNER JOIN his_estrucra ON empleado.ternro = his_estructura.ternro " & _
             " INNER JOIN fer_estr ON fer_estr.tenro = his_estructura.tenro " & _
             " INNER JOIN feriado ON fer_estr.ferinro = feriado.ferinro " & _
             " WHERE empleado.ternro = " & buliq_empleado!ternro & _
             " AND feriado.tipferinro = 2" & _
             " AND feriado.ferifecha >= " & ConvFecha(Desde) & _
             " AND feriado.ferifecha < " & ConvFecha(Hasta)
    OpenRecordset StrSql, rs_feriados
    
    Do While Not rs_feriados.EOF
        If Weekday(rs_feriados!ferifecha) > 1 Then
            DiasH = DiasH - 1
        End If
        
        ' Siguiente Feriado
        rs_feriados.MoveNext
    Loop
    
    
    ' cierro todo y libero
    If rs_pais.State = adStateOpen Then rs_pais.Close
    If rs_feriados.State = adStateOpen Then rs_feriados.Close
        
    Set rs_feriados = Nothing
    Set rs_pais = Nothing

End Sub



Public Sub bus_Cotmon1()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtenci�n de la cotizaci�n de la Moneda. gcotmon1.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroMon As Long          'moneda.monnro
Dim Opcion As Long          '1-fecha desde del proceso
                            '2-fecha hasta del proceso
                            '3-fecha de pago del proceso
                            '4-fecha desde del periodo
                            '5-fecha hasta del periodo
                            '6-today

'Dim Param_cur As New ADODB.Recordset
Dim rs_cotizacion As New ADODB.Recordset
Dim AFecha As Date

    Bien = False
    Valor = 0
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroMon = Arr_Programa(NroProg).Auxint1
        Opcion = Arr_Programa(NroProg).Auxint2
    Else
        Exit Sub
    End If

Select Case Opcion
Case 1: 'fecha desde del proceso
    AFecha = buliq_proceso!profecini
Case 2: 'fecha hasta del proceso
    AFecha = buliq_proceso!profecfin
Case 3: 'fecha de pago del proceso
    AFecha = buliq_proceso!profecpago
Case 4: 'fecha desde del periodo
    AFecha = buliq_periodo!pliqdesde
Case 5: 'fecha hasta del periodo
    AFecha = buliq_periodo!pliqhasta
Case 6: 'Today
    AFecha = Date
Case Else
    'tipo de fecha no valido
End Select


    ' Busco la ultima cotizaci�n
    StrSql = "SELECT * FROM cotizamon WHERE monnro = " & NroMon & _
             " AND cotfecha <= " & AFecha & _
             " ORDER BY fecha"
    OpenRecordset StrSql, rs_cotizacion
    
    If Not rs_cotizacion.EOF Then
        rs_cotizacion.MoveLast
        
'        If Opcion = 1 Then
'            valor = rs_cotizacion!cotvalororigen
'        Else
'            valor = rs_cotizacion!cotvalorinternac
'        End If
        
        Valor = rs_cotizacion!cotvalororigen
    Else
        Valor = 0
        Bien = False
    End If

' Cierro y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_cotizacion.State = adStateOpen Then rs_cotizacion.Close
Set rs_cotizacion = Nothing
'Set Param_cur = Nothing

End Sub



Public Sub bus_Tcotpa0()
' ---------------------------------------------------------------------------------------------
' Descripcion: Evalua para todos los conceptos de dicho tipo la b�squeda asociada
'              segun el alcance de resoluci�n para el empleado y Sumariza todos
'              los resultados. gtcotpa0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroTipoCon As Long  'tipconcep.tconnro
Dim NroTipoPar As Long  'tipopar.tpanro
Dim val As Single
Dim fec As Date
Dim Ok As Boolean

Dim param_cur       As New ADODB.Recordset
Dim rs_buf_concepto As New ADODB.Recordset
Dim rs_Concepto     As New ADODB.Recordset
Dim rs_Formula      As New ADODB.Recordset
Dim rs_cft          As New ADODB.Recordset
Dim rs_Con_For_Tpa  As New ADODB.Recordset
Dim rs_Programa     As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Evalua para todos los conceptos de dicho tipo la b�squeda asociada "
        Flog.writeline Espacios(Tabulador * 4) & "segun el alcance de resoluci�n para el empleado y Sumariza todos los resultados."
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroTipoCon = Arr_Programa(NroProg).Auxint1
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de Concepto (tipconcep.tconnro) " & NroTipoCon
        End If
        
        NroTipoPar = Arr_Programa(NroProg).Auxint2
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Tipo de Parametro (tipopar.tpanro) " & NroTipoPar
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros cargados "
        End If
        
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If
    
    
    Bien = True
    Valor = 0
     
    StrSql = "SELECT * FROM concepto WHERE tconnro = " & CStr(NroTipoCon)
    OpenRecordset StrSql, rs_buf_concepto
    
    Do While Not rs_buf_concepto.EOF
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Concepto " & rs_buf_concepto!concnro
        End If
    
        'posicionarse en cada pasada en el buliq a liquidar
        StrSql = "SELECT * FROM concepto " & _
                 " INNER JOIN formula ON concepto.fornro = formula.fornro " & _
                 " INNER JOIN for_tpa ON formula.fornro = for_tpa.fornro " & _
                 " WHERE concnro = " & rs_buf_concepto!concnro & _
                 " AND for_tpa.tpanro = " & NroTipoPar & _
                 " ORDER BY tpa_fornro, for_tpa.ftorden "
        OpenRecordset StrSql, rs_Concepto
        
        ' Resoluci�n de los Par�metros de la F�rmula del Concepto
        Do While Not rs_Concepto.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 5) & "Parametro " & rs_Concepto!tpanro
            End If
        
            StrSql = "SELECT * FROM cft_segun WHERE concnro = " & rs_Concepto!concnro & _
                     " AND tpanro = & rs_concepto.tpanro " & _
                     " AND nivel = 1 " & _
                     " AND origen = " & NroGrupo
            OpenRecordset StrSql, rs_cft
            
            If rs_cft.EOF Then
                If rs_cft.State = adStateOpen Then rs_cft.Close
                    StrSql = "SELECT * FROM cft_segun WHERE concnro = " & rs_Concepto!concnro & _
                             " AND tpanro = & rs_concepto.tpanro " & _
                             " AND nivel = 2 "
                OpenRecordset StrSql, rs_cft
                
                If rs_cft.EOF Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 6) & "Sin alcance "
                    End If
                    GoTo SiguienteConcepto
                End If
            End If
            
            StrSql = "SELECT * FROM con_for_tpa WHERE concnro = " & rs_Concepto!concnro & _
                     " AND tpanro = & rs_concepto.tpanro " & _
                     " AND nivel = " & rs_cft!Nivel & _
                     " AND selecc = " & rs_cft!Selecc
            OpenRecordset StrSql, rs_Con_For_Tpa
                    
            If rs_Con_For_Tpa.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 6) & "Sin alcance "
                End If
                GoTo SiguienteConcepto
            End If
                
            StrSql = "SELECT * FROM programa WHERE programa.prognro = " & rs_Con_For_Tpa!Prognro
            OpenRecordset StrSql, rs_Programa
            
            If rs_Programa.EOF Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 6) & "Sin busqueda asociada "
                End If
                GoTo SiguienteConcepto
            End If
                
            val = 0
            ' Si es autom�tico y la busqueda est� marcada como que puede usar cache,
            ' --> verificar el cache del empleado
            If rs_Con_For_Tpa!auto And rs_Programa!Progcache Then
                If objCache.EsSimboloDefinido(CStr(rs_Programa!Prognro)) Then
                    val = objCache.Valor(CStr(rs_Programa!Prognro))
                    Ok = True
                Else    ' Busqueda Automatica, Primera vez
                    Call EjecutarBusqueda(15, rs_Concepto!concnro, rs_Concepto!tpanro, val, fec, Ok)
                    ' insertar en el cache del empleado
                    Call objCache.Insertar_Simbolo(CStr(rs_Programa!Prognro), val)
                End If
            Else    ' Busqueda NO automatica, Novedades: Buscar
                Call EjecutarBusqueda(15, rs_Concepto!concnro, rs_Concepto!tpanro, val, fec, Ok)
            End If
                            
            If Ok Then  ' Se obtuvo el parametro satisfactoriamente
                Valor = Valor + val
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 6) & "Se obtuvo el parametro satisfactoriamente. Valor = " & val
                End If
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 6) & "Busqueda no exitosa "
                End If
            End If
              
SiguienteConcepto:
            rs_Concepto.MoveNext
        Loop
        
        rs_buf_concepto.MoveNext
    Loop


' Cierro y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_buf_concepto.State = adStateOpen Then rs_buf_concepto.Close
If rs_Concepto.State = adStateOpen Then rs_Concepto.Close
If rs_Formula.State = adStateOpen Then rs_Formula.Close
If rs_cft.State = adStateOpen Then rs_cft.Close
If rs_Con_For_Tpa.State = adStateOpen Then rs_Con_For_Tpa.Close
If rs_Programa.State = adStateOpen Then rs_Programa.Close

'Set Param_cur = Nothing
Set rs_buf_concepto = Nothing
Set rs_Concepto = Nothing
Set rs_Formula = Nothing
Set rs_cft = Nothing
Set rs_Con_For_Tpa = Nothing
Set rs_Programa = Nothing

End Sub

Public Sub Bus_NovGegi(ByVal Prognro As Long, ByVal concepto As Long, ByVal tpanro As Long, ByVal fecha_inicio As Date, ByVal fecha_fin As Date, ByVal Grupo As Long, ByRef Ok As Boolean, ByRef val As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de novedad a Nivel General/Grupal/Individual. novgegi.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_NovEstr As New ADODB.Recordset
Dim rs_NovGral As New ADODB.Recordset
Dim rs_His_Estructura As New ADODB.Recordset
Dim rs_HisNovEmp As New ADODB.Recordset
Dim rs_HisNovEstr As New ADODB.Recordset
Dim rs_HisNovGral As New ADODB.Recordset

Dim Encontro As Boolean
Dim Aux_Encontro As Boolean
Dim Vigencia_Activa As Boolean

'Dim ChequeaFirmas As Boolean
Dim Firmado As Boolean
Dim rs_firmas As New ADODB.Recordset

    Encontro = False
    Ok = False
    
If Not Encontro And (Prognro = 3 Or Prognro = 12 Or Prognro = 13 Or Prognro = 14) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Novedad Individual, por empleado "
    End If
    'Case 3, 12, 13, 14:
        StrSql = "SELECT * FROM novemp WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND empleado = " & buliq_cabliq!Empleado & _
                 " AND ((nevigencia = -1 " & _
                 " AND nedesde <= " & ConvFecha(fecha_fin) & _
                 " AND (nehasta >= " & ConvFecha(fecha_inicio) & _
                 " OR nehasta is null )) " & _
                 " OR nevigencia = 0)" & _
                 " ORDER BY nevigencia, nedesde, nehasta "
        OpenRecordset StrSql, rs_NovEmp
        
        val = 0
        Do While Not rs_NovEmp.EOF
            If FirmaActiva5 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                         " AND cysfircodext = '" & rs_NovEmp!nenro & "' and cystipnro = 5"
                OpenRecordset StrSql, rs_firmas
                If rs_firmas.EOF Then
                    Firmado = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                    End If
                Else
                    Firmado = True
                End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If

        If Firmado Then
            If CBool(rs_NovEmp!nevigencia) Then
                Vigencia_Activa = True
                If Not EsNulo(rs_NovEmp!nehasta) Then
                    If (rs_NovEmp!nehasta < fecha_inicio) Or (fecha_fin < rs_NovEmp!nedesde) Then
                        'Exit Sub
                        Vigencia_Activa = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " INACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " ACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    End If
                Else
                    If (fecha_fin < rs_NovEmp!nedesde) Then
                        'Exit Sub
                        Vigencia_Activa = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    End If
                End If
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Novedad sin vigencia con valor " & rs_NovEmp!nevalor
                End If
            End If
            
            If Vigencia_Activa Or Not CBool(rs_NovEmp!nevigencia) Then
                val = val + rs_NovEmp!nevalor
            End If
'            If Not EsNulo(rs_NovEmp!neretro) Then
'                fec = rs_NovEmp!neretro
'            End If
            
            If Not (EsNulo(rs_NovEmp!nepliqdesde) Or EsNulo(rs_NovEmp!nepliqhasta)) Then
                If rs_NovEmp!nepliqdesde <> 0 And rs_NovEmp!nepliqhasta <> 0 Then
                    Retroactivo = True
                    pliqdesde = rs_NovEmp!nepliqdesde
                    pliqhasta = rs_NovEmp!nepliqhasta
                Else
                    Retroactivo = False
                End If
            Else
                Retroactivo = False
            End If
            
            Ok = True
            Encontro = True
            If guarda_nov Then
                'Crea la novedad hist�rica
                StrSql = "SELECT * FROM hisnovemp WHERE" & _
                         " concnro = " & rs_NovEmp!concnro & _
                         " AND tpanro =" & rs_NovEmp!tpanro & _
                         " AND empleado =" & rs_NovEmp!Empleado & _
                         " AND pronro =" & buliq_proceso!pronro
                OpenRecordset StrSql, rs_HisNovEmp
                
                If rs_HisNovEmp.EOF Then
                    ' Inserta
                    If Not EsNulo(rs_NovEmp!nepliqdesde) Then
                        StrSql = "INSERT INTO hisnovemp ("
                        StrSql = StrSql & "concnro,nevalor,tpanro,empleado,pronro,fechis,nepliqdesde,nepliqhasta,nevigencia"
                        If Not EsNulo(rs_NovEmp!nedesde) Then
                            StrSql = StrSql & ",nedesde"
                        End If
                        If Not EsNulo(rs_NovEmp!nehasta) Then
                            StrSql = StrSql & ",nehasta"
                        End If
                        StrSql = StrSql & ") VALUES (" & rs_NovEmp!concnro
                        StrSql = StrSql & "," & rs_NovEmp!nevalor
                        StrSql = StrSql & "," & rs_NovEmp!tpanro
                        StrSql = StrSql & "," & rs_NovEmp!Empleado
                        StrSql = StrSql & "," & buliq_proceso!pronro
                        StrSql = StrSql & "," & ConvFecha(Date)
                        StrSql = StrSql & "," & rs_NovEmp!nepliqdesde
                        StrSql = StrSql & "," & rs_NovEmp!nepliqhasta
                        StrSql = StrSql & "," & CInt(rs_NovEmp!nevigencia)
                        If Not EsNulo(rs_NovEmp!nedesde) Then
                            StrSql = StrSql & "," & ConvFecha(rs_NovEmp!nedesde)
                        End If
                        If Not EsNulo(rs_NovEmp!nehasta) Then
                            StrSql = StrSql & "," & ConvFecha(rs_NovEmp!nehasta)
                        End If
                        StrSql = StrSql & " )"
                    Else
                        StrSql = "INSERT INTO hisnovemp ("
                        StrSql = StrSql & "concnro,nevalor,tpanro,empleado,pronro,fechis,nevigencia"
                        If Not EsNulo(rs_NovEmp!nedesde) Then
                            StrSql = StrSql & ",nedesde"
                        End If
                        If Not EsNulo(rs_NovEmp!nehasta) Then
                            StrSql = StrSql & ",nehasta"
                        End If
                        StrSql = StrSql & ") VALUES (" & rs_NovEmp!concnro
                        StrSql = StrSql & "," & rs_NovEmp!nevalor
                        StrSql = StrSql & "," & rs_NovEmp!tpanro
                        StrSql = StrSql & "," & rs_NovEmp!Empleado
                        StrSql = StrSql & "," & buliq_proceso!pronro
                        StrSql = StrSql & "," & ConvFecha(Date)
                        StrSql = StrSql & "," & CInt(rs_NovEmp!nevigencia)
                        If Not EsNulo(rs_NovEmp!nedesde) Then
                            StrSql = StrSql & "," & ConvFecha(rs_NovEmp!nedesde)
                        End If
                        If Not EsNulo(rs_NovEmp!nehasta) Then
                            StrSql = StrSql & "," & ConvFecha(rs_NovEmp!nehasta)
                        End If
                        StrSql = StrSql & " )"
                    End If
                    
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If 'If Not rs_HisNovEmp.EOF Then
            End If 'If guarda_nov Then
        End If 'If Firmado Then
        
        rs_NovEmp.MoveNext
    Loop
    If Not Encontro Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad Individual "
        End If
    End If
End If

If Not Encontro And (Prognro = 2 Or Prognro = 11 Or Prognro = 12 Or Prognro = 14) Then
'Case 2, 11, 12, 14:
        'buscar por Estructura
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Por estructura "
        End If
        StrSql = "SELECT * FROM novestr WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND ((ntevigencia = -1 " & _
                 " AND ntedesde <= " & ConvFecha(fecha_fin) & _
                 " AND (ntehasta >= " & ConvFecha(fecha_inicio) & " " & _
                 " OR ntehasta is null)) " & _
                 " OR ntevigencia = 0) " & _
                 " ORDER BY ntevigencia, ntedesde, ntehasta "
        OpenRecordset StrSql, rs_NovEstr
        
        Ok = False
        Encontro = False
        Aux_Encontro = False
        If rs_NovEstr.EOF Then
            Firmado = False
        End If
        val = 0
        Do While Not rs_NovEstr.EOF 'And Not Encontro
            If FirmaActiva15 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovEstr!estrnro & "' and cystipnro = 15"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                        End If
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
            
            If Firmado Then
                If CBool(rs_NovEstr!ntevigencia) Then
                    Vigencia_Activa = True
                    If Not EsNulo(rs_NovEstr!ntehasta) Then
                        If (rs_NovEstr!ntehasta < fecha_inicio) Or (fecha_fin < rs_NovEstr!ntedesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta " & rs_NovEstr!ntehasta & " INACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta " & rs_NovEstr!ntehasta & " ACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        End If
                    Else
                        If (fecha_fin < rs_NovEstr!ntedesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ntedesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEstr!ntevalor
                            End If
                        End If
                    End If
                End If
                
                If Vigencia_Activa Or Not CBool(rs_NovEstr!ntevigencia) Then
                    Encontro = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "busco que el empleado tenga la estructura " & rs_NovEstr!estrnro & " activa"
                    End If
                    
                    'busco que el empleado tenga esa estructura activa
                    StrSql = " SELECT tenro, estrnro FROM his_estructura " & _
                             " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                             " estrnro =" & rs_NovEstr!estrnro & _
                             " AND (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND " & _
                             " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_His_Estructura
                    If Not rs_His_Estructura.EOF Then
                        val = val + rs_NovEstr!ntevalor
'                        If Not EsNulo(rs_NovEstr!nteretro) Then
'                            fec = rs_NovEstr!nteretro
'                        End If

                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Estructura " & rs_NovEstr!estrnro & " activa"
                        End If
                        Ok = True 'esta faltando retroactividad
                        Encontro = True
                        Aux_Encontro = True
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Estructura " & rs_NovEstr!estrnro & " No activa"
                        End If
                    End If
                End If
            End If 'firmado
           
            If Encontro And guarda_nov Then
                'Crea la novedad hist�rica
                StrSql = "SELECT * FROM hisnovgru WHERE" & _
                         " concnro = " & rs_NovEstr!concnro & _
                         " AND tpanro =" & rs_NovEstr!tpanro & _
                         " AND grunro =" & rs_NovEstr!estrnro & _
                         " AND pronro =" & buliq_proceso!pronro
                OpenRecordset StrSql, rs_HisNovEstr
                
                If rs_HisNovEstr.EOF Then
                    ' Inserta
                    StrSql = "INSERT INTO hisnovgru (" & _
                             "concnro,ngvalor,tpanro,grunro,pronro,fechis" & _
                             ") VALUES (" & rs_NovEstr!concnro & _
                             "," & rs_NovEstr!ntevalor & _
                             "," & rs_NovEstr!tpanro & _
                             "," & rs_NovEstr!estrnro & _
                             "," & buliq_proceso!pronro & _
                             "," & ConvFecha(Date) & _
                             " )"
                             ' esta faltando retroactividad
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If 'If Not rs_HisNovEstr.EOF Then
            End If 'If guarda_nov Then
            
            rs_NovEstr.MoveNext
        Loop
        Encontro = Aux_Encontro
        If Not Encontro Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad por Estructura"
            End If
        End If
End If

If Not Encontro And (Prognro = 1 Or Prognro = 11 Or Prognro = 12 Or Prognro = 13) Then
'Case 1, 11, 12, 13:

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Global "
    End If
    
' FGZ - 09/02/2004
    If objCache_NovGlobales.EsSimboloDefinido(CStr(concepto & "-" & tpanro)) Then
        val = objCache_NovGlobales.Valor(CStr(concepto & "-" & tpanro))
        
        Ok = True 'esta faltando retroactividad
        Encontro = True
        If guarda_nov Then
            'Crea la novedad hist�rica
            StrSql = "SELECT * FROM hisnovgral WHERE"
            StrSql = StrSql & " concnro = " & rs_NovGral!concnro
            StrSql = StrSql & " AND tpanro =" & rs_NovGral!tpanro
            StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
            OpenRecordset StrSql, rs_HisNovGral
            
            If rs_HisNovGral.EOF Then
                ' Inserta
                StrSql = "INSERT INTO hisnovgral ("
                StrSql = StrSql & "concnro,ngravalor,tpanro,pronro,fechis"
                StrSql = StrSql & ") VALUES (" & rs_NovGral!concnro
                StrSql = StrSql & "," & rs_NovGral!ngravalor
                StrSql = StrSql & "," & rs_NovGral!tpanro
                StrSql = StrSql & "," & buliq_proceso!pronro
                StrSql = StrSql & "," & ConvFecha(Date)
                StrSql = StrSql & " )"
                         ' esta faltando retroactividad
                objConn.Execute StrSql, , adExecuteNoRecords
            End If 'If Not rs_HisNovGral.EOF Then
        End If 'If guarda_nov Then
        
    Else
        StrSql = "SELECT * FROM novgral WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND ((ngravigencia = -1 " & _
                 " AND ngradesde < " & ConvFecha(fecha_fin) & " " & _
                 " AND (ngrahasta >= " & ConvFecha(fecha_inicio) & " " & _
                 " OR ngrahasta is null)) " & _
                 " OR ngravigencia = 0) " & _
                 " ORDER BY ngravigencia, ngradesde, ngrahasta "
        OpenRecordset StrSql, rs_NovGral
           
        Do While Not rs_NovGral.EOF
            If FirmaActiva19 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovGral!ngranro & "' and cystipnro = 19"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                        End If
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
        
            If Firmado Then
                If CBool(rs_NovGral!ngravigencia) Then
                    Vigencia_Activa = True
                    If Not EsNulo(rs_NovGral!ngrahasta) Then
                        If (rs_NovGral!ngrahasta < fecha_inicio) Or (fecha_fin < rs_NovGral!ngradesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " INACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " ACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        End If
                    Else
                        If (fecha_fin < rs_NovGral!ngradesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado INACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado ACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        End If
                    End If
                End If
                
                If Vigencia_Activa Or Not CBool(rs_NovGral!ngravigencia) Then
                    val = val + rs_NovGral!ngravalor
                    
'                    If Not EsNulo(rs_NovGral!ngraretro) Then
'                        fec = rs_NovGral!ngraretro
'                    End If
'                    If rs_NovGral!ngrapliqdesde <> 0 And rs_NovGral!ngrapliqhasta <> 0 Then
'                        Retroactivo = True
'                        pliqdesde = rs_NovGral!ngrapliqdesde
'                        pliqhasta = rs_NovGral!ngrapliqhasta

'                    Else
'                        Retroactivo = False
'                    End If
                    
                    
                    Ok = True 'esta faltando retroactividad
                    Encontro = True
                    If guarda_nov Then
                        'Crea la novedad hist�rica
                        StrSql = "SELECT * FROM hisnovgral WHERE"
                        StrSql = StrSql & " concnro = " & rs_NovGral!concnro
                        StrSql = StrSql & " AND tpanro =" & rs_NovGral!tpanro
                        StrSql = StrSql & " AND pronro =" & buliq_proceso!pronro
                        OpenRecordset StrSql, rs_HisNovGral
                        
                        If rs_HisNovGral.EOF Then
                            ' Inserta
                            StrSql = "INSERT INTO hisnovgral ("
                            StrSql = StrSql & "concnro,ngravalor,tpanro,pronro,fechis"
                            StrSql = StrSql & ") VALUES (" & rs_NovGral!concnro
                            StrSql = StrSql & "," & rs_NovGral!ngravalor
                            StrSql = StrSql & "," & rs_NovGral!tpanro
                            StrSql = StrSql & "," & buliq_proceso!pronro
                            StrSql = StrSql & "," & ConvFecha(Date)
                            StrSql = StrSql & " )"
                                     ' esta faltando retroactividad
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If 'If Not rs_HisNovGral.EOF Then
                    End If 'If guarda_nov Then
                End If
            End If 'If Firmado Then
            
            rs_NovGral.MoveNext
        Loop
        
        If Encontro Then
            'inserto la novedad en el cache
            Call objCache_NovGlobales.Insertar_Simbolo(CStr(concepto & "-" & tpanro & "0"), val)
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad global"
            End If
        End If
    End If
End If

' Libero y cierro
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
Set rs_NovEmp = Nothing

If rs_NovEstr.State = adStateOpen Then rs_NovEstr.Close
Set rs_NovEstr = Nothing

If rs_NovGral.State = adStateOpen Then rs_NovGral.Close
Set rs_NovGral = Nothing

'Historicos
If rs_HisNovEmp.State = adStateOpen Then rs_HisNovEmp.Close
Set rs_HisNovEmp = Nothing

If rs_HisNovEstr.State = adStateOpen Then rs_HisNovEstr.Close
Set rs_HisNovEstr = Nothing

If rs_HisNovGral.State = adStateOpen Then rs_HisNovGral.Close
Set rs_HisNovGral = Nothing

End Sub

Public Sub Bus_NovGegi_Nuevo(ByVal Prognro As Long, ByVal concepto As Long, ByVal tpanro As Long, ByVal fecha_inicio As Date, ByVal fecha_fin As Date, ByVal Grupo As Long, ByRef Ok As Boolean, ByRef val As Single, ByRef TipoNovedad As String)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de novedad a Nivel General/Grupal/Individual. novgegi.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_NovEstr As New ADODB.Recordset
Dim rs_NovGral As New ADODB.Recordset

Dim rs_HisNovEmp As New ADODB.Recordset
Dim rs_HisNovEstr As New ADODB.Recordset
Dim rs_HisNovGral As New ADODB.Recordset

Dim Encontro As Boolean
    
Dim ChequeaFirmas As Boolean
Dim Firmado As Boolean
Dim rs_firmas As New ADODB.Recordset

    Encontro = False
    Ok = False
    
BuscarIndividual:
If Not Encontro Then
    'Case 3, 12, 13, 14:
'    Call RevisarFirmas(5, ChequeaFirmas)
        StrSql = "SELECT * FROM novemp WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND empleado = " & buliq_cabliq!Empleado
        OpenRecordset StrSql, rs_NovEmp
        
        If Not rs_NovEmp.EOF Then
            If FirmaActiva5 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                         " AND cysfircodext = '" & rs_NovEmp!nenro & "' and cystipnro = 5"
                OpenRecordset StrSql, rs_firmas
                If rs_firmas.EOF Then
                    Firmado = False
                Else
                    Firmado = True
                End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
        Else
            Firmado = False
        End If

        If Firmado Then
            If CBool(rs_NovEmp!nevigencia) Then
                If (rs_NovEmp!nehasta < fecha_inicio) Or (fecha_fin < rs_NovEmp!nedesde) Then
                    GoTo BuscarPorEstructura
                    'Exit Sub
                End If
            End If
            
            val = rs_NovEmp!nevalor
'            If Not EsNulo(rs_NovEmp!neretro) Then
'                fec = rs_NovEmp!neretro
'            End If
            
            If Not (EsNulo(rs_NovEmp!nepliqdesde) Or EsNulo(rs_NovEmp!nepliqhasta)) Then
                If rs_NovEmp!nepliqdesde <> 0 And rs_NovEmp!nepliqhasta <> 0 Then
                    Retroactivo = True
                    pliqdesde = rs_NovEmp!nepliqdesde
                    pliqhasta = rs_NovEmp!nepliqhasta
    
                Else
                    Retroactivo = False
                End If
            Else
                Retroactivo = False
            End If
            
            Ok = True
            Encontro = True
            TipoNovedad = "Individual"
            
            If guarda_nov Then
                'Crea la novedad hist�rica
                StrSql = "SELECT * FROM hisnovemp WHERE" & _
                         " concnro = " & rs_NovEmp!concnro & _
                         " AND tpanro =" & rs_NovEmp!tpanro & _
                         " AND empleado =" & rs_NovEmp!Empleado & _
                         " AND pronro =" & buliq_proceso!pronro
                OpenRecordset StrSql, rs_HisNovEmp
                
                If rs_HisNovEmp.EOF Then
                    ' Inserta
                    If Not EsNulo(rs_NovEmp!nepliqdesde) Then
                        StrSql = "INSERT INTO hisnovemp (" & _
                                 "concnro,nevalor,nevigencia,nedesde,nehsta,tpanro,empleado,pronro,fechis,nepliqdesde,nepliqhasta" & _
                                 ") VALUES (" & rs_NovEmp!concnro & _
                                 "," & rs_NovEmp!nevalor & _
                                 "," & rs_NovEmp!tpanro & _
                                 "," & rs_NovEmp!Empleado & _
                                 "," & buliq_proceso!pronro & _
                                 "," & ConvFecha(Date) & _
                                 "," & rs_NovEmp!nepliqdesde & _
                                 "," & rs_NovEmp!nepliqhasta & _
                                 " )"
                    Else
                        StrSql = "INSERT INTO hisnovemp (" & _
                                 "concnro,nevalor,tpanro,empleado,pronro,fechis" & _
                                 ") VALUES (" & rs_NovEmp!concnro & _
                                 "," & rs_NovEmp!nevalor & _
                                 "," & rs_NovEmp!tpanro & _
                                 "," & rs_NovEmp!Empleado & _
                                 "," & buliq_proceso!pronro & _
                                 "," & ConvFecha(Date) & _
                                 " )"
                    End If
                    
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If 'If Not rs_HisNovEmp.EOF Then
            End If 'If guarda_nov Then
        End If 'If Firmado Then
End If

BuscarPorEstructura:
If Not Encontro Then
'Case 2, 11, 12, 14:
'buscar por Estructura
'    Call RevisarFirmas(15, ChequeaFirmas)
        StrSql = "SELECT * FROM novestr WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND estrnro = " & Grupo
        OpenRecordset StrSql, rs_NovEstr
        
        If Not rs_NovEstr.EOF Then
            If FirmaActiva15 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovEstr!nanro & "' and cystipnro = 15"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
        Else
            Firmado = False
        End If
            
        If Firmado Then
            If CBool(rs_NovEstr!ntevigencia) Then
                If (rs_NovEstr!ntehasta < fecha_inicio) Or (fecha_fin < rs_NovEstr!ntedesde) Then
                    GoTo BuscarGlobal
                    'Exit Sub
                End If
            End If
            
            val = rs_NovEstr!ntevalor
'            If Not EsNulo(rs_NovEstr!nteretro) Then
'                fec = rs_NovEstr!nteretro
'            End If
            
            'If rs_NovEstr!nepliqdesde <> 0 And rs_NovEstr!nepliqhasta <> 0 Then
            '    retroactivo = True
            'Else
            '    retroactivo = False
            'End If
            '
            'pliqdesde = rs_NovEstr!nepliqdesde
            'pliqhasta = rs_NovEstr!nepliqhasta
            Ok = True 'esta faltando retroactividad
            Encontro = True
            TipoNovedad = "Estructura"
            
            If guarda_nov Then
                'Crea la novedad hist�rica
                StrSql = "SELECT * FROM hisnovgru WHERE" & _
                         " concnro = " & rs_NovEstr!concnro & _
                         " AND tpanro =" & rs_NovEstr!tpanro & _
                         " AND grunro =" & rs_NovEstr!estrnro & _
                         " AND pronro =" & buliq_proceso!pronro
                OpenRecordset StrSql, rs_HisNovEstr
                
                If Not rs_HisNovEstr.EOF Then
                    ' Inserta
                    StrSql = "INSERT INTO hisnovgru (" & _
                             "concnro,ngvalor,tpanro,grunro,pronro,fechis" & _
                             ") VALUES (" & rs_NovEstr!concnro & _
                             "," & rs_NovEstr!ntevalor & _
                             "," & rs_NovEstr!tpanro & _
                             "," & rs_NovEstr!estrnro & _
                             "," & rs_NovEstr!pronro & _
                             "," & ConvFecha(Date) & _
                             " )"
                             ' esta faltando retroactividad
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If 'If Not rs_HisNovEstr.EOF Then
            End If 'If guarda_nov Then
    End If 'If Firmado Then
End If

BuscarGlobal:
If Not Encontro Then
'Case 1, 11, 12, 13:
' buscar general
'    Call RevisarFirmas(19, ChequeaFirmas)
    StrSql = "SELECT * FROM novgral WHERE " & _
             " concnro = " & concepto & _
             " AND tpanro = " & tpanro
    OpenRecordset StrSql, rs_NovGral
    
    If Not rs_NovGral.EOF Then
        If FirmaActiva19 Then
            '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                         " AND cysfircodext = '" & rs_NovGral!nanro & "' and cystipnro = 19"
                OpenRecordset StrSql, rs_firmas
                If rs_firmas.EOF Then
                    Firmado = False
                Else
                    Firmado = True
                End If
            If rs_firmas.State = adStateOpen Then rs_firmas.Close
        Else
            Firmado = True
        End If
    Else
        Firmado = False
    End If
    
    If Firmado Then
            If CBool(rs_NovGral!ngravigencia) Then
                If (rs_NovGral!ngrahasta < fecha_inicio) Or (fecha_fin < rs_NovGral!ngradesde) Then
                    Exit Sub
                End If
            End If
            
            val = rs_NovGral!ngravalor
'            If Not EsNulo(rs_NovGral!ngraretro) Then
'                fec = rs_NovGral!ngraretro
'            End If
'            If rs_NovGral!ngrapliqdesde <> 0 And rs_NovGral!ngrapliqhasta <> 0 Then
'                Retroactivo = True
'                pliqdesde = rs_NovGral!ngrapliqdesde
'                pliqhasta = rs_NovGral!ngrapliqhasta
'            Else
'                Retroactivo = False
'            End If
            
            
            Ok = True 'esta faltando retroactividad
            Encontro = True
            TipoNovedad = "Global"
            
            If guarda_nov Then
                'Crea la novedad hist�rica
                StrSql = "SELECT * FROM hisnovgral WHERE" & _
                         " concnro = " & rs_NovGral!concnro & _
                         " AND tpanro =" & rs_NovGral!tpanro & _
                         " AND pronro =" & buliq_proceso!pronro
                OpenRecordset StrSql, rs_HisNovGral
                
                If Not rs_HisNovGral.EOF Then
                    ' Inserta
                    StrSql = "INSERT INTO hisnovgral (" & _
                             "concnro,ngravalor,tpanro,pronro,fechis" & _
                             ") VALUES (" & rs_NovGral!concnro & _
                             "," & rs_NovGral!ngravalor & _
                             "," & rs_NovGral!tpanro & _
                             "," & rs_NovGral!pronro & _
                             "," & ConvFecha(Date) & _
                             " )"
                             ' esta faltando retroactividad
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If 'If Not rs_HisNovGral.EOF Then
            End If 'If guarda_nov Then
    End If 'If Firmado Then
End If

' Libero y cierro
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
Set rs_NovEmp = Nothing

If rs_NovEstr.State = adStateOpen Then rs_NovEstr.Close
Set rs_NovEstr = Nothing

If rs_NovGral.State = adStateOpen Then rs_NovGral.Close
Set rs_NovGral = Nothing

'Historicos
If rs_HisNovEmp.State = adStateOpen Then rs_HisNovEmp.Close
Set rs_HisNovEmp = Nothing

If rs_HisNovEstr.State = adStateOpen Then rs_HisNovEstr.Close
Set rs_HisNovEstr = Nothing

If rs_HisNovGral.State = adStateOpen Then rs_HisNovGral.Close
Set rs_HisNovGral = Nothing

End Sub


Public Sub Bus_NovGegiHis(ByVal Prognro As Long, ByVal concepto As Long, ByVal tpanro As Long, ByVal fecha_inicio As Date, ByVal fecha_fin As Date, ByVal Grupo As Long, ByRef Ok As Boolean, ByRef val As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de novedad a Nivel General/Grupal/Individual. novgegi.p del Historico
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_NovEstr As New ADODB.Recordset
Dim rs_NovGral As New ADODB.Recordset
Dim rs_His_Estructura As New ADODB.Recordset
Dim rs_HisNovEmp As New ADODB.Recordset
Dim rs_HisNovEstr As New ADODB.Recordset
Dim rs_HisNovGral As New ADODB.Recordset

Dim Encontro As Boolean
Dim Aux_Encontro As Boolean
Dim Vigencia_Activa As Boolean

Dim Firmado As Boolean
Dim rs_firmas As New ADODB.Recordset

    Encontro = False
    Ok = False
    
If Not Encontro And (Prognro = 3 Or Prognro = 12 Or Prognro = 13 Or Prognro = 14) Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Novedad Individual, por empleado "
    End If
    'Case 3, 12, 13, 14:
        StrSql = "SELECT * FROM hisnovemp WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND pronro = " & buliq_proceso!pronro & _
                 " AND empleado = " & buliq_cabliq!Empleado & _
                 " AND ((nevigencia = -1 " & _
                 " AND nedesde <= " & ConvFecha(fecha_fin) & _
                 " AND (nehasta >= " & ConvFecha(fecha_inicio) & _
                 " OR nehasta is null )) " & _
                 " OR nevigencia = 0)" & _
                 " ORDER BY nevigencia, nedesde, nehasta "
        OpenRecordset StrSql, rs_NovEmp
        
        val = 0
        Do While Not rs_NovEmp.EOF
            If FirmaActiva5 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                         " AND cysfircodext = '" & rs_NovEmp!nenro & "' and cystipnro = 5"
                OpenRecordset StrSql, rs_firmas
                If rs_firmas.EOF Then
                    Firmado = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                    End If
                Else
                    Firmado = True
                End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If

        If Firmado Then
            If CBool(rs_NovEmp!nevigencia) Then
                Vigencia_Activa = True
                If Not EsNulo(rs_NovEmp!nehasta) Then
                    If (rs_NovEmp!nehasta < fecha_inicio) Or (fecha_fin < rs_NovEmp!nedesde) Then
                        'Exit Sub
                        Vigencia_Activa = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " INACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta " & rs_NovEmp!nehasta & " ACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    End If
                Else
                    If (fecha_fin < rs_NovEmp!nedesde) Then
                        'Exit Sub
                        Vigencia_Activa = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEmp!nedesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEmp!nevalor
                        End If
                    End If
                End If
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Novedad sin vigencia con valor " & rs_NovEmp!nevalor
                End If
            End If
            
            If Vigencia_Activa Or Not CBool(rs_NovEmp!nevigencia) Then
                val = val + rs_NovEmp!nevalor
            End If
            
            If Not (EsNulo(rs_NovEmp!nepliqdesde) Or EsNulo(rs_NovEmp!nepliqhasta)) Then
                If rs_NovEmp!nepliqdesde <> 0 And rs_NovEmp!nepliqhasta <> 0 Then
                    Retroactivo = True
                    pliqdesde = rs_NovEmp!nepliqdesde
                    pliqhasta = rs_NovEmp!nepliqhasta
                Else
                    Retroactivo = False
                End If
            Else
                Retroactivo = False
            End If
            
            Ok = True
            Encontro = True
        End If 'If Firmado Then
        
        rs_NovEmp.MoveNext
    Loop
    If Not Encontro Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad Individual "
        End If
    End If
End If

If Not Encontro And (Prognro = 2 Or Prognro = 11 Or Prognro = 12 Or Prognro = 14) Then
'Case 2, 11, 12, 14:
        'buscar por Estructura
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Por estructura "
        End If
        StrSql = "SELECT * FROM hisnovgru WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND pronro = " & buliq_proceso!pronro & _
                 " AND ((ngvigencia = -1 " & _
                 " AND ngdesde <= " & ConvFecha(fecha_fin) & _
                 " AND (nghasta >= " & ConvFecha(fecha_inicio) & " " & _
                 " OR nghasta is null)) " & _
                 " OR ngvigencia = 0) " & _
                 " ORDER BY ngvigencia, ngdesde, nghasta "
        OpenRecordset StrSql, rs_NovEstr
        
        Ok = False
        Encontro = False
        Aux_Encontro = False
        If rs_NovEstr.EOF Then
            Firmado = False
        End If
        val = 0
        Do While Not rs_NovEstr.EOF 'And Not Encontro
            If FirmaActiva15 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovEstr!estrnro & "' and cystipnro = 15"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                        End If
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
            
            If Firmado Then
                If CBool(rs_NovEstr!ngvigencia) Then
                    Vigencia_Activa = True
                    If Not EsNulo(rs_NovEstr!nghasta) Then
                        If (rs_NovEstr!nghasta < fecha_inicio) Or (fecha_fin < rs_NovEstr!ngdesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ngdesde & " Hasta " & rs_NovEstr!nghasta & " INACTIVA con valor " & rs_NovEstr!ngvalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ngdesde & " Hasta " & rs_NovEstr!nghasta & " ACTIVA con valor " & rs_NovEstr!ngvalor
                            End If
                        End If
                    Else
                        If (fecha_fin < rs_NovEstr!ngdesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ngdesde & " Hasta indeterminado INACTIVA con valor " & rs_NovEstr!ngvalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovEstr!ngdesde & " Hasta indeterminado ACTIVA con valor " & rs_NovEstr!ngvalor
                            End If
                        End If
                    End If
                End If
                
                If Vigencia_Activa Or Not CBool(rs_NovEstr!ngvigencia) Then
                    Encontro = False
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "busco que el empleado tenga la estructura " & rs_NovEstr!estrnro & " activa"
                    End If
                    
                    'busco que el empleado tenga esa estructura activa
                    StrSql = " SELECT tenro, estrnro FROM his_estructura " & _
                             " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                             " estrnro =" & rs_NovEstr!estrnro & _
                             " AND (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND " & _
                             " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
                    OpenRecordset StrSql, rs_His_Estructura
                    If Not rs_His_Estructura.EOF Then
                        val = val + rs_NovEstr!ngvalor

                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Estructura " & rs_NovEstr!estrnro & " activa"
                        End If
                        Ok = True 'esta faltando retroactividad
                        Encontro = True
                        Aux_Encontro = True
                    Else
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "Estructura " & rs_NovEstr!estrnro & " No activa"
                        End If
                    End If
                End If
            End If 'firmado
           
            rs_NovEstr.MoveNext
        Loop
        Encontro = Aux_Encontro
        If Not Encontro Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad por Estructura"
            End If
        End If
End If

If Not Encontro And (Prognro = 1 Or Prognro = 11 Or Prognro = 12 Or Prognro = 13) Then
'Case 1, 11, 12, 13:

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Global "
    End If
    
' FGZ - 09/02/2004
    If objCache_NovGlobales.EsSimboloDefinido(CStr(concepto & "-" & tpanro)) Then
        val = objCache_NovGlobales.Valor(CStr(concepto & "-" & tpanro))
        
        Ok = True 'esta faltando retroactividad
        Encontro = True
    Else
        StrSql = "SELECT * FROM novgral WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND ((ngravigencia = -1 " & _
                 " AND ngradesde < " & ConvFecha(fecha_fin) & " " & _
                 " AND (ngrahasta >= " & ConvFecha(fecha_inicio) & " " & _
                 " OR ngrahasta is null)) " & _
                 " OR ngravigencia = 0) " & _
                 " ORDER BY ngravigencia, ngradesde, ngrahasta "
        OpenRecordset StrSql, rs_NovGral
           
        Do While Not rs_NovGral.EOF
            If FirmaActiva19 Then
                '/* Verificar si esta en el NIVEL FINAL DE FIRMA */
                    StrSql = "select * from cysfirmas where cysfirfin = -1 " & _
                             " AND cysfircodext = '" & rs_NovGral!ngranro & "' and cystipnro = 19"
                    OpenRecordset StrSql, rs_firmas
                    If rs_firmas.EOF Then
                        Firmado = False
                        If CBool(USA_DEBUG) Then
                            Flog.writeline Espacios(Tabulador * 4) & "NIVEL FINAL DE FIRMA No Activo "
                        End If
                    Else
                        Firmado = True
                    End If
                If rs_firmas.State = adStateOpen Then rs_firmas.Close
            Else
                Firmado = True
            End If
        
            If Firmado Then
                If CBool(rs_NovGral!ngravigencia) Then
                    Vigencia_Activa = True
                    If Not EsNulo(rs_NovGral!ngrahasta) Then
                        If (rs_NovGral!ngrahasta < fecha_inicio) Or (fecha_fin < rs_NovGral!ngradesde) Then
                            'Exit Sub
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " INACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta " & rs_NovGral!ngrahasta & " ACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        End If
                    Else
                        If (fecha_fin < rs_NovGral!ngradesde) Then
                            Vigencia_Activa = False
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado INACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        Else
                            If CBool(USA_DEBUG) Then
                                Flog.writeline Espacios(Tabulador * 4) & "Novedad con vigencia desde " & rs_NovGral!ngradesde & " Hasta indeterminado ACTIVA con valor " & rs_NovGral!ngravalor
                            End If
                        End If
                    End If
                End If
                
                If Vigencia_Activa Or Not CBool(rs_NovGral!ngravigencia) Then
                    val = val + rs_NovGral!ngravalor
                    
                    Ok = True 'esta faltando retroactividad
                    Encontro = True
                End If
            End If 'If Firmado Then
            
            rs_NovGral.MoveNext
        Loop
        
        If Encontro Then
            'inserto la novedad en el cache
            Call objCache_NovGlobales.Insertar_Simbolo(CStr(concepto & "-" & tpanro & "0"), val)
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Novedad global"
            End If
        End If
    End If
End If

' Libero y cierro
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
Set rs_NovEmp = Nothing

If rs_NovEstr.State = adStateOpen Then rs_NovEstr.Close
Set rs_NovEstr = Nothing

If rs_NovGral.State = adStateOpen Then rs_NovGral.Close
Set rs_NovGral = Nothing

'Historicos
If rs_HisNovEmp.State = adStateOpen Then rs_HisNovEmp.Close
Set rs_HisNovEmp = Nothing

If rs_HisNovEstr.State = adStateOpen Then rs_HisNovEstr.Close
Set rs_HisNovEstr = Nothing

If rs_HisNovGral.State = adStateOpen Then rs_HisNovGral.Close
Set rs_HisNovGral = Nothing

End Sub



Public Sub Bus_NovGegiHis_old(ByVal Prognro As Long, ByVal concepto As Long, ByVal tpanro As Long, ByVal fecha_inicio As Date, ByVal fecha_fin As Date, ByVal Grupo As Long, ByRef Ok As Boolean, ByRef val As Single)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de novedad Historicas a Nivel General/Grupal/Individual.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_NovEstr As New ADODB.Recordset
Dim rs_NovGral As New ADODB.Recordset
Dim rs_His_Estructura As New ADODB.Recordset
Dim Encontro As Boolean
    

Encontro = False
Ok = False
    
If Not Encontro And (Prognro = 3 Or Prognro = 12 Or Prognro = 13 Or Prognro = 14) Then
    'Case 3, 12, 13, 14:
        StrSql = "SELECT * FROM hisnovemp WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND pronro =" & buliq_proceso!tpanro & _
                 " AND empleado = " & buliq_cabliq!Empleado
        OpenRecordset StrSql, rs_NovEmp
        
        If Not rs_NovEmp.EOF Then
            val = rs_NovEmp!nevalor
'            If Not EsNulo(rs_NovEmp!neretro) Then
'                fec = rs_NovEmp!neretro
'            End If
            
            If rs_NovEmp!nepliqdesde <> 0 And rs_NovEmp!nepliqhasta <> 0 Then
                Retroactivo = True
                pliqdesde = rs_NovEmp!nepliqdesde
                pliqhasta = rs_NovEmp!nepliqhasta
            Else
                Retroactivo = False
            End If
        End If
End If

If Not Encontro And (Prognro = 2 Or Prognro = 11 Or Prognro = 12 Or Prognro = 14) Then
'Case 2, 11, 12, 14:
'/* buscar por GRUPO */
'        StrSql = "SELECT * FROM hisnovgrup WHERE " & _
'                 " concnro = " & concepto & _
'                 " AND tpanro = " & tpanro & _
'                 " AND grunro = " & Grupo & _
'                 " AND pronro =" & buliq_proceso!tpanro
        
        StrSql = "SELECT * FROM hisnovgrup WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND pronro =" & buliq_proceso!tpanro
        OpenRecordset StrSql, rs_NovEstr
        
        If Not rs_NovEstr.EOF Then
            'busco que el empleado tenga esa estructura activa
            StrSql = " SELECT tenro, estrnro FROM his_estructura " & _
                     " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                     " (htetdesde <= " & ConvFecha(Empleado_Fecha_Fin) & ") AND " & _
                     " ((" & ConvFecha(Empleado_Fecha_Fin) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_His_Estructura
            If Not rs_His_Estructura.EOF Then
                val = rs_NovEstr!ntevalor
'                If Not EsNulo(rs_NovEstr!nteretro) Then
'                    fec = rs_NovEstr!nteretro
'                End If
                
                
                Ok = True 'esta faltando retroactividad
                Encontro = True
            Else
                Ok = False 'esta faltando retroactividad
                Encontro = False
            End If
        End If
End If

If Not Encontro And (Prognro = 1 Or Prognro = 11 Or Prognro = 12 Or Prognro = 13) Then
'Case 1, 11, 12, 13:
' buscar general
        StrSql = "SELECT * FROM hisnovgral WHERE " & _
                 " concnro = " & concepto & _
                 " AND tpanro = " & tpanro & _
                 " AND pronro =" & buliq_proceso!tpanro
        OpenRecordset StrSql, rs_NovGral
        
        If Not rs_NovGral.EOF Then
            val = rs_NovGral!ngravalor
'            If Not EsNulo(rs_NovGral!ngraretro) Then
'                fec = rs_NovGral!ngraretro
'            End If
           
            Ok = True 'esta faltando retroactividad
            Encontro = True
        End If
End If

' Libero y cierro
If rs_NovEmp.State = adStateOpen Then rs_NovEmp.Close
Set rs_NovEmp = Nothing

If rs_NovEstr.State = adStateOpen Then rs_NovEstr.Close
Set rs_NovEstr = Nothing

If rs_NovGral.State = adStateOpen Then rs_NovGral.Close
Set rs_NovGral = Nothing

End Sub



Public Sub bus_Grilla(ByVal tipoBus As Long, ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala. Ggrilla0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroGrilla As Long               ' cabgrilla.cgrnro
Dim Cero_No_Encuentra As Boolean    ' True - Devuelve 0 si no encuentra?
Dim Operacion As Long               ' 1 - Sumatoria
                                    ' 2 - Maximo
                                    ' 3 - Promedio
                                    ' 4 - Promedio sin 0
                                    ' 5 - Minimo
                                    ' 6 - Primer valor no vacio desde abajo
                                    ' 7 - Primer valor no vacio desde arriba
                                    
Dim Acumulativa As Boolean          ' True
                                    ' False
Dim Valor_Grilla(10) As Boolean     ' Elemento de una coordenada de una grilla

'Dim Param_cur As New ADODB.Recordset
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & " Busca el valor de una escala "
    End If
    
    
    ' Obtener los parametros de la Busqueda
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'If Param_cur.State = adStateOpen Then Param_cur.Close
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroGrilla = Arr_Programa(NroProg).Auxint1
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Nro de escala a buscar  " & NroGrilla
        End If
        
        Acumulativa = CBool(Arr_Programa(NroProg).Auxlog1)
        Operacion = Arr_Programa(NroProg).Auxint2
        Cero_No_Encuentra = CBool(Arr_Programa(NroProg).Auxlog2)
        ' valores de la grilla
        Valor_Grilla(1) = CBool(Arr_Programa(NroProg).Auxlog3)
        Valor_Grilla(2) = CBool(Arr_Programa(NroProg).Auxlog4)
        Valor_Grilla(3) = CBool(Arr_Programa(NroProg).Auxlog5)
        Valor_Grilla(4) = CBool(Arr_Programa(NroProg).Auxlog6)
        Valor_Grilla(5) = CBool(Arr_Programa(NroProg).Auxlog7)
        Valor_Grilla(6) = CBool(Arr_Programa(NroProg).Auxlog8)
        Valor_Grilla(7) = CBool(Arr_Programa(NroProg).Auxlog9)
        Valor_Grilla(8) = CBool(Arr_Programa(NroProg).Auxlog10)
        Valor_Grilla(9) = CBool(Arr_Programa(NroProg).Auxlog11)
        Valor_Grilla(10) = CBool(Arr_Programa(NroProg).Auxlog12)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If

'SEGUN CONDICION LLAMA A bus_grilla0 o bus_grilla1
If Not Acumulativa Then
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Grilla No Acumulativa "
        Flog.writeline Espacios(Tabulador * 4) & "Procediiento Grilla0 "
    End If
    
    Call bus_grilla0(NroGrilla, Cero_No_Encuentra, Valor_Grilla, Operacion, Acumulativa, tipoBus, concnro, prog)
Else
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Grilla Acumulativa "
        Flog.writeline Espacios(Tabulador * 4) & "Procediiento Grilla1 "
    End If

    'Call bus_grilla1(NroGrilla, Cero_No_Encuentra, Valor_Grilla, Operacion, Acumulativa)
    Call bus_grilla1(NroGrilla, Cero_No_Encuentra, Valor_Grilla, Operacion, Acumulativa, tipoBus, concnro, prog)
End If

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing
End Sub


Private Sub CargarValoresdelaGrilla(ByVal rs As ADODB.Recordset, ByRef Arreglo)
' ---------------------------------------------------------------------------------------------
' Descripcion: Llena un arreglo con los valores de los registros de ValGrilla.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim i As Integer

rs.MoveFirst
i = 1
    Do While Not rs.EOF
        If Not EsNulo(rs!vgrvalor) Then
            Arreglo(i) = rs!vgrvalor
            i = i + 1
        End If
        
        rs.MoveNext
    Loop

End Sub


Public Sub bus_grilla0(ByVal NroGrilla As Long, ByVal Cero_No_Encuentra As Boolean, ByVal Valor_Grilla, ByVal Operacion As Integer, ByVal Acumulativa As Boolean, ByVal tipoBus As Long, ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala. Ggrilla0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Single     ' para alojar los valores de:  valgrilla.val(i)

Dim TipoBase As Long
Dim TipoBaseVariable As Long

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim rs_Busqueda As New ADODB.Recordset

Dim NroBusqueda As Long
Dim TipoBusqueda As Long
Dim Encontro As Boolean

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

'    ' Buscar el tipo Base de la antiguedad
'    StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
'             " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
'             " WHERE tipoprog.tprogbase = 4"
'    OpenRecordset StrSql, rs_tbase
'
'    If Not rs_tbase.EOF Then
'        TipoBase = rs_tbase!tprogbase
'    End If
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    'FGZ - 22/01/2004
    TipoBaseVariable = 15
    
    Continuar = True
    ant = 1
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Grilla de " & rs_cabgrilla!cgrdimension & "dimensiones "
    End If
    
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

' FGZ - 22/01/2004
'Parametros Variables
' busco que parametro es el parametro del concepto
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "busco que parametro es el parametro del concepto "
    End If

    Continuar = True
    pvar = 1
    Do While (pvar <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case pvar
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBaseVariable = rs_tbase!tprogbase Then
                    Continuar = False
                    pvariable = True
                Else
                    pvar = pvar + 1
                End If
            End If
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBaseVariable = rs_tbase!tprogbase Then
                    Continuar = False
                    pvariable = True
                Else
                    pvar = pvar + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBaseVariable = rs_tbase!tprogbase Then
                    Continuar = False
                    pvariable = True
                Else
                    pvar = pvar + 1
                End If
            End If
        Case 4:
            If rs_cabgrilla!grparnro_4 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        Case 5:
            If rs_cabgrilla!grparnro_5 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        End Select
    Loop


    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Resuelvo los indices de la grilla segun las busquedas por cada dimension"
    End If
   
    For j = 1 To rs_cabgrilla!cgrdimension
        ' rs_cabgrilla!grparnro_x tiene el nro de programa

        Select Case j
        Case 1:
            NroBusqueda = rs_cabgrilla!grparnro_1
        Case 2:
            NroBusqueda = rs_cabgrilla!grparnro_2
        Case 3:
            NroBusqueda = rs_cabgrilla!grparnro_3
        Case 4:
            NroBusqueda = rs_cabgrilla!grparnro_4
        Case 5:
            NroBusqueda = rs_cabgrilla!grparnro_5
        Case Else
        End Select
        
        StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroBusqueda)
        OpenRecordset StrSql, rs_Busqueda
    
        If Not rs_Busqueda.EOF Then
            TipoBusqueda = rs_Busqueda!Tprognro
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Error. no se encontro el nro de busqueda "
            End If
            'Error. no se encontro el nro de busqueda
            Exit Sub
        End If
        
        Call EjecutarBusqueda(TipoBusqueda, concnro, NroBusqueda, Valor, fec, False)
        Parametros(j) = Valor
    Next j

    If Not antig Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No busca antiguedad "
        End If
    
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
'        For j = 1 To rs_cabgrilla!cgrdimension
'            StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
'        Next j
' FGZ - 22/01/2004
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant And j <> pvar Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            Else
                If pvariable Then
                    StrSql = StrSql & " AND vgrcoor_" & j & "<= " & Parametros(j)
                End If
            End If
                    
            'StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
        Next j
        If pvariable Then
            StrSql = StrSql & " ORDER BY vgrcoor_" & pvar & " DESC "
        End If
        
        OpenRecordset StrSql, rs_valgrilla
    
        If Not rs_valgrilla.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Cargo los Valores de la Grilla "
            End If
            Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
            
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco el valor segun la operacion "
            End If
            Call BusValor(Operacion, Valor_Grilla, grilla_val, Valor)
        Else
            If Cero_No_Encuentra Then
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontr� valor en grilla "
                    Flog.writeline Espacios(Tabulador * 4) & "Esta configurado que retorne cero si no lo encuentra "
                End If
            
                 Valor = 0
                 Bien = True
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontr� valor en grilla "
                    Flog.writeline Espacios(Tabulador * 4) & "Retorna Falso "
                End If
                Bien = False
            End If
       End If
    Else 'Antig
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busca antiguedad "
        End If
    
        If Not Cero_No_Encuentra Then
            Bien = False
        Else
            Bien = True
            Valor = 0
        End If
    
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco la primera antiguedad de la escala menor a la del empleado "
            Flog.writeline Espacios(Tabulador * 4) & "de abajo hacia arriba "
        End If
    
        'Busco la primera antiguedad de la escala menor a la del empleado
        ' de abajo hacia arriba
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
            StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        
        Encontro = False
        
        Do While Not rs_valgrilla.EOF And Not Encontro
            'Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
            '--------------------------
            Select Case ant
            Case 1:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 2:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 3:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 4:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 5:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            End Select
            '--------------------------
            
'            If Parametros(ant) >= rs_valgrilla(ant) Then
'                Call BusValor(Operacion, Valor_Grilla, grilla_val, valor)
'                Bien = True
'            End If
            
            rs_valgrilla.MoveNext
        Loop
        If CBool(USA_DEBUG) Then
            If Encontro Then
                Flog.writeline Espacios(Tabulador * 4) & "Valor encontrado "
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Valor no encontrado "
                If Cero_No_Encuentra Then
                    Flog.writeline Espacios(Tabulador * 4) & "Esta configurado que retorne cero si no lo encuentra "
                Else
                    Flog.writeline Espacios(Tabulador * 4) & "No Esta configurado que retorne cero si no lo encuentra. Retorna Falso "
                End If
            End If
        End If
    End If
    
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Private Sub BusValor(ByVal Op As Integer, ByVal valorGrilla, ByVal valgrilla, ByRef Valor As Single)
Dim cant As Integer
Dim Continuar As Boolean
Dim i As Integer


Select Case Op
Case 1:     'Sumatoria
    Valor = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            Valor = Valor + valgrilla(i)
        End If
    Next i

Case 2:     'Maximo
    Valor = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            If valgrilla(i) > Valor Then
                Valor = valgrilla(i)
            End If
        End If
    Next i
    
Case 3:     'Promedio
    Valor = 0
    cant = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            Valor = Valor + valgrilla(i)
            cant = cant + 1
        End If
    Next i

    If cant <> 0 Then
        Valor = Valor / cant
    End If

Case 4:     'Promedio sin cero
    Valor = 0
    cant = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            If valgrilla(i) <> 0 Then
                Valor = Valor + valgrilla(i)
                cant = cant + 1
            End If
        End If
    Next i

    If cant <> 0 Then
        Valor = Valor / cant
    End If

Case 5:     'Minimo
    Valor = 0
    For i = 1 To 10
        If valorGrilla(i) Then
            If Valor = 0 Or valgrilla(i) < Valor Then
                Valor = valgrilla(i)
            End If
        End If
    Next i

Case 6:     'Primer valor no vacio desde el primero
    Valor = 0
    i = 1
    Continuar = True
    Do While i <= 10 And Continuar
        If valorGrilla(i) Then
            If valgrilla(i) <> 0 Then
                Valor = valgrilla(i)
                Continuar = False
            End If
        End If
        i = i + 1
    Loop

Case 7:     'Primer valor no vacio desde el ultimo
    Valor = 0
    i = 10
    Continuar = True
    Do While i >= 0 And Continuar
        If valorGrilla(i) Then
            If valgrilla(i) <> 0 Then
                Valor = valgrilla(i)
                Continuar = False
            End If
        End If
        i = i - 1
    Loop

End Select

End Sub

Public Sub bus_grilla1(ByVal NroGrilla As Long, ByVal Cero_No_Encuentra As Boolean, ByVal Valor_Grilla, ByVal Operacion As Integer, ByVal Acumulativa As Boolean, ByVal tipoBus As Long, ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala. Ggrilla0.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Single     ' para alojar los valores de:  valgrilla.val(i)

Dim TipoBase As Long

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim rs_Busqueda As New ADODB.Recordset

Dim NroBusqueda As Long
Dim TipoBusqueda As Long
Dim Encontro As Boolean

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    ' Buscar el tipo Base de la antiguedad
    StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
             " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
             " WHERE tipoprog.tprogbase = 4"
    OpenRecordset StrSql, rs_tbase

    If Not rs_tbase.EOF Then
        TipoBase = rs_tbase!tprogbase
    End If

    'El tipo Base de la antiguedad
    TipoBase = 4
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase,tipoprog.tprognro FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                    antig = True
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop

    
'Parametros Variables
' busco que parametro es el par�metro del concepto
    Continuar = True
    pvar = 1
    Do While (pvar <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case pvar
        Case 1:
            If rs_cabgrilla!grparnro_1 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        Case 2:
            If rs_cabgrilla!grparnro_2 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        Case 3:
            If rs_cabgrilla!grparnro_1 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        Case 4:
            If rs_cabgrilla!grparnro_4 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        Case 5:
            If rs_cabgrilla!grparnro_5 = 15 Then 'si es el param. variable
                Continuar = False
                pvariable = True
            Else
                pvar = pvar + 1
            End If
        End Select
    Loop
    
    
    For j = 1 To rs_cabgrilla!cgrdimension
        ' rs_cabgrilla!grparnro_x tiene el nro de programa

        Select Case j
        Case 1:
            NroBusqueda = rs_cabgrilla!grparnro_1
        Case 2:
            NroBusqueda = rs_cabgrilla!grparnro_2
        Case 3:
            NroBusqueda = rs_cabgrilla!grparnro_3
        Case 4:
            NroBusqueda = rs_cabgrilla!grparnro_4
        Case 5:
            NroBusqueda = rs_cabgrilla!grparnro_5
        Case Else
        End Select
        
        StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroBusqueda)
        OpenRecordset StrSql, rs_Busqueda
    
        If Not rs_Busqueda.EOF Then
            TipoBusqueda = rs_Busqueda!Tprognro
        Else
            'Error. no se encontro el nro de busqueda
            Exit Sub
        End If
        
        Call EjecutarBusqueda(TipoBusqueda, concnro, NroBusqueda, Valor, fec, False)
        'Call bus_Estructura(rs_cabgrilla!grparnro_1)
        'Call EjecutarBusqueda(rs_Programa!tprognro, rs_Conceptos!concnro, rs_Con_For_Tpa!prognro, val, fec, ok)
        'Call CalcularBusqueda(NroGrilla, j, NroBusqueda, valor, tipoBus, concnro)
        Parametros(j) = Valor
    Next j




    If Not antig Then
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant And j <> pvar Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            Else
                If pvariable Then
                    StrSql = StrSql & " AND vgrcoor_" & j & "<= " & Parametros(j)
                End If
            End If
                    
            'StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
        Next j
        If pvariable Then
            StrSql = StrSql & " ORDER BY vgrcoor_" & pvar & " DESC "
        End If
        OpenRecordset StrSql, rs_valgrilla
    
        If Not rs_valgrilla.EOF Then
            Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
    
            Call BusValor(Operacion, Valor_Grilla, grilla_val, Valor)
        Else
            If Cero_No_Encuentra Then
                 Valor = 0
                 Bien = True
            Else
                Bien = False
            End If
       End If
    Else 'Antig
        If Not Cero_No_Encuentra Then
            Bien = False
        Else
            Bien = True
            Valor = 0
        End If
    
        'Busco la primera antiguedad de la escala menor a la del empleado
        ' de abajo hacia arriba
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
            StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        
        Encontro = False
        
        Do While Not rs_valgrilla.EOF And Not Encontro
            Select Case ant
            Case 1:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 2:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 3:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 4:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            Case 5:
                If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                     If rs_valgrilla!vgrvalor <> 0 Then
                        Valor = rs_valgrilla!vgrvalor
                        Encontro = True
                     End If
                End If
            End Select
            
            rs_valgrilla.MoveNext
        Loop
    End If
    
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub bus_grilla1_old(ByVal NroGrilla As Long, ByVal Cero_No_Encuentra As Boolean, ByVal Valor_Grilla, ByVal Operacion As Integer, ByVal Acumulativa As Boolean)
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala. Ggrilla1.p
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim Desde As Single
Dim aux As Single

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_grparametro As New ADODB.Recordset

Dim NroBusqueda As Long
'
'If Not Bien Then
'    ' En la generacion de la busqueda atomica se selecciona una busqueda en la escala acumulativa
'    ' y en una escala no se encontr� el parametro.
'    ' 15 - Parametro del concepto
'    Exit Sub
'End If
'
'    StrSql = "SELECT * FROM cabgrilla " & _
'             " WHERE cabgrilla.cgrnro = " & NroGrilla
'    OpenRecordset StrSql, rs_cabgrilla
'
'
'    For j = 1 To rs_cabgrilla!cgrdimension
'        ' rs_cabgrilla!grparnro_1 tiene el nro de programa
'
'        Select Case j
'        Case 1:
'            NroBusqueda = rs_cabgrilla!grparnro_1
'        Case 2:
'            NroBusqueda = rs_cabgrilla!grparnro_2
'        Case 3:
'            NroBusqueda = rs_cabgrilla!grparnro_3
'        Case 4:
'            NroBusqueda = rs_cabgrilla!grparnro_4
'        Case 5:
'            NroBusqueda = rs_cabgrilla!grparnro_5
'        Case Else
'        End Select
'        Call CalcularBusqueda(NroGrilla, j, NroBusqueda, valor)
'        Parametros(j) = valor
'    Next j
'
'
'    Bien = False
'
'    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
'    For j = 1 To rs_cabgrilla!cgrdimension
'        If j <> pvar Then
'            StrSql = StrSql & "AND vgrcoor_" & j & "= " & Parametros(j)
'        Else
'            StrSql = StrSql & "AND vgrcoor_" & j & "<= " & Parametros(j)
'        End If
'    Next j
'
'    StrSql = StrSql & " ORDER BY vgrcoor_" & pvar & "DESC"
'    OpenRecordset StrSql, rs_valgrilla
'
'    Do While Not rs_valgrilla.EOF
'        Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
'
'        ' setea en vgrcoor_pvar el valor de lo que seria en progress:   vgrcoor[pvar]
'        Select Case pvar
'        Case 1:
'            vgrcoor_pvar = rs_valgrilla!vgrcoor_1
'        Case 2:
'            vgrcoor_pvar = rs_valgrilla!vgrcoor_2
'        Case 3:
'            vgrcoor_pvar = rs_valgrilla!vgrcoor_3
'        Case 4:
'            vgrcoor_pvar = rs_valgrilla!vgrcoor_4
'        Case 5:
'            vgrcoor_pvar = rs_valgrilla!vgrcoor_5
'        End Select
'
'        If vgrcoor_pvar <= Parametros(pvar) Then
'            Call BusValor(Operacion, Valor_Grilla, grilla_val, Aux)
'            valor = valor + ((vgrcoor_pvar - Desde) * Aux) / 100
'            If vgrcoor_pvar = Parametros(pvar) Then
'                Bien = True
'                Exit Sub
'            End If
'        Else
'            Call BusValor(Operacion, Valor_Grilla, grilla_val, Aux)
'            valor = valor + ((Parametros(pvar) - Desde) * Aux) / 100
'            Bien = True
'            Exit Sub
'        End If
'
'       rs_valgrilla.MoveNext
'    Loop
'
'
'' Cierro todo y libero
'If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
'If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close
'If rs_grparametro.State = adStateOpen Then rs_grparametro.Close
'
'Set rs_cabgrilla = Nothing
'Set rs_valgrilla = Nothing
'Set rs_grparametro = Nothing
End Sub


Public Sub CalcularBusqueda_OLD(ByVal NroGrilla As Long, ByVal Posicion As Integer, ByVal paranro As Long, ByRef valorcor As Integer)
' -----------------------------------------------------------------------------------
' Descripcion: liqgrbus.p. Programa de busqueda de parametro de la grilla
' Autor: FGZ
' Fecha: 31/07/2003
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim antdia As Integer
Dim antmeses As Integer
Dim antanio As Integer
Dim diashab As Integer

Dim Aux_cgrtpg As String
Dim ret As Integer

Dim rs_domicilio As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_grupo As New ADODB.Recordset
Dim rs_TipoEmpleador As New ADODB.Recordset


Select Case paranro
Case 1: '
    valorcor = buliq_empleado!catnro
Case 2: '
    valorcor = buliq_empleado!convnro
Case 3: '
    valorcor = buliq_empleado!tenro
Case 4: '
    valorcor = buliq_empleado!reghornro
Case 5: '
    ret = Antiguedad(antdia, antmeses, antanio, diashab)
    If ret = 0 Then ' sin errores
        valorcor = (antanio * 12) + antmeses
    End If
Case 6: '
    valorcor = buliq_empleado!puenro
Case 7: '
    valorcor = buliq_empleado!sucursal
Case 8: '
    valorcor = buliq_empleado!actnro
Case 9: '
    valorcor = buliq_empleado!gernro
Case 10: '
    valorcor = buliq_empleado!folinro
Case 11: '
    valorcor = buliq_empleado!mobrnro
Case 12: '
    valorcor = buliq_empleado!celulanro
Case 13: '
    valorcor = buliq_empleado!lineanro
Case 14: '
    valorcor = buliq_empleado!maqnro
Case 15: '
    ' WF_TPA
Case 16: 'Sector
    valorcor = buliq_empleado!secnro
Case 17: 'Concepto
    valorcor = Buliq_Concepto(Concepto_Actual).concnro
Case 18: 'Zona del domicilio de default de la sucursal

    StrSql = "SELECT zona.zonanro FROM tercero " & _
    " INNER JOIN cabdom ON tercero.ternro = cabdom.ternro " & _
    " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
    " INNER JOIN zona ON detdom.zonanro = zona.zonanro " & _
    " WHERE tercero.ternro = " & buliq_empleado!sucursal
    OpenRecordset StrSql, rs_domicilio

    If rs_domicilio.EOF Then
        Exit Sub
    Else
        valorcor = rs_domicilio!zonanro
    End If

Case 19: 'Zona del domicilio de la sucursal o del domicilio de trabajo del empleado: ESPECIAL MSD

    ' verificar el nro de sucursal, trato especial
    StrSql = "SELECT sucursal.succod, sucursal.ternro FROM sucursal " & _
    " INNER JOIN tercero ON tercero.ternro = sucursal.ternro " & _
    " WHERE sucursal.ternro = " & buliq_empleado!sucursal
    OpenRecordset StrSql, rs_Sucursal

    If rs_Sucursal.EOF Then
        Exit Sub
    Else
        If rs_Sucursal!succod <> 4 Then
            StrSql = "SELECT zona.zonanro FROM tercero " & _
            " INNER JOIN cabdom ON tercero.ternro = cabdom.ternro " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
            " INNER JOIN zona ON detdom.zonanro = zona.zonanro " & _
            " WHERE tercero.ternro = " & rs_Sucursal!ternro & _
            " AND cabdom.domdefault = -1 AND cabdom.tidonro = 10 "
            OpenRecordset StrSql, rs_domicilio

            If rs_domicilio.EOF Then
                Exit Sub
            Else
                valorcor = rs_domicilio!zonanro
            End If
        Else
            StrSql = "SELECT zona.zonanro FROM tercero " & _
            " INNER JOIN cabdom ON tercero.ternro = cabdom.ternro " & _
            " INNER JOIN detdom ON detdom.domnro = cabdom.domnro " & _
            " INNER JOIN zona ON detdom.zonanro = zona.zonanro " & _
            " WHERE tercero.ternro = " & buliq_empleado!sucursal & _
            " AND cabdom.tidonro = 2 "
            OpenRecordset StrSql, rs_domicilio

            If rs_domicilio.EOF Then
                Exit Sub
            Else
                valorcor = rs_domicilio!zonanro
            End If
        End If
    End If

Case 20: '
    valorcor = buliq_empleado!ccosnro

Case 21: '
    valorcor = buliq_empleado!tprocnro

Case 27: 'Grupos
    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla
    If Not rs_cabgrilla.EOF Then

        ' seteo el nombre del parametro de la tabla cabgrilla segun posicion
        Select Case Posicion
        Case 1:
            Aux_cgrtpg = rs_cabgrilla!cgrtpg__1
        Case 2:
            Aux_cgrtpg = rs_cabgrilla!cgrtpg__2
        Case 3:
            Aux_cgrtpg = rs_cabgrilla!cgrtpg__3
        Case 4:
            Aux_cgrtpg = rs_cabgrilla!cgrtpg__4
        Case 5:
            Aux_cgrtpg = rs_cabgrilla!cgrtpg__5
        Case Else
        End Select

        StrSql = "SELECT * FROM tip_grup_ter " & _
                 " WHERE tgnro = " & Aux_cgrtpg & _
                 " AND ternro =" & buliq_empleado!ternro

        OpenRecordset StrSql, rs_grupo

        If Not rs_grupo.EOF Then
            valorcor = rs_grupo!grunro
        Else
            valorcor = 0
        End If
    Else
        StrSql = "SELECT * FROM tip_grup_ter " & _
                 " WHERE tgnro = " & Posicion & _
                 " AND ternro =" & buliq_empleado!ternro
        OpenRecordset StrSql, rs_grupo

        If Not rs_grupo.EOF Then
            valorcor = rs_grupo!grunro
        Else
            valorcor = 0
        End If
    End If

Case 40: 'Indemnizacion
    Call bus_Antiguedad("INDEMNIZACON", buliq_periodo!pliqhasta, antdia, antmeses, antanio, diashab)
    valorcor = (antanio * 12) + antmeses

Case 41: 'Vacaciones
    Call bus_Antiguedad("VACACIONES", buliq_periodo!pliqhasta, antdia, antmeses, antanio, diashab)
    valorcor = (antanio * 12) + antmeses

Case 42: 'Obra Social por Defecto
    valorcor = buliq_empleado!osocialxdef

Case 43: 'Obra Social elegida
    valorcor = buliq_empleado!osocial

Case 44: 'Sindicato
    valorcor = buliq_empleado!gremio

Case 45: 'Departamento
    valorcor = buliq_empleado!depnro

Case 46: 'Direccion
    valorcor = buliq_empleado!dirnro

Case 48: 'Todos Concepto
    valorcor = Buliq_Concepto(Concepto_Actual).concnro

Case 49: 'Tipo de Empleador

        StrSql = "SELECT * FROM empresa " & _
                 " INNER JOIN tipempdor ON empresa.tipempnro = tipempdor.tipempnro " & _
                 " WHERE empnro = " & NroEmp

        OpenRecordset StrSql, rs_TipoEmpleador

        If Not rs_TipoEmpleador.EOF Then
            valorcor = rs_TipoEmpleador!tipempnro
        End If

Case 50: 'Situacion de Revista
    valorcor = buliq_empleado!srnro

Case 51: 'Condicion
    valorcor = buliq_empleado!csijpcod

Case Else

End Select
End Sub

'                            NroGrilla , j, NroBusqueda, valor, tipoBus, concnro
Public Sub CalcularBusqueda(ByVal NroGrilla As Long, ByVal Posicion As Integer, ByVal prog As Long, ByRef valorcor As Single, ByVal tipoBus As Long, ByVal concnro As Long)
' -----------------------------------------------------------------------------------
' Descripcion: liqgrbus.p. Programa de busqueda de parametro de la grilla
' Autor: FGZ
' Fecha: 31/07/2003
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
' FGZ - 09/02/2004
If objCache.EsSimboloDefinido(CStr(prog)) Then
     Valor = objCache.Valor(CStr(prog))
     Bien = True
 Else
     Call EjecutarBusqueda(tipoBus, concnro, prog, Valor, fec, Bien)
     If Not Bien Then
         If HACE_TRAZA Then
             Call InsertarTraza(NroCab, concnro, prog, "Error Busqueda de Escala", 0)
         End If
     End If
     ' insertar en el cache del empleado
     Call objCache.Insertar_Simbolo(CStr(prog), Valor)
 End If

If Bien Then
    valorcor = Valor
End If


End Sub

Public Function Antiguedad(ByRef dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer, ByRef DiasHabiles As Integer) As Integer
' -----------------------------------------------------------------------------------
' Descripcion: Antigued.p. Calcula la antiguedad al dia de hoy de un empleado en :
'               dias h�biles(si es menor que un a�o) o en dias, meses y a�os en caso contrario.
'               Retorna 0 si no hubo error y <> 0 en caso contrario
' Autor: FGZ
' Fecha: 31/07/2003
' Ultima Modificacion:
' -----------------------------------------------------------------------------------
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer
Dim NombreCampo As String
Dim rs_Fases As New ADODB.Recordset

DiasHabiles = 0

StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro
OpenRecordset StrSql, rs_Fases

If rs_Fases.EOF Then
    ' ERROR. El empleado no tiene fecha de alta en fases
    Antiguedad = 1
    Exit Function
Else
        fecalta = rs_Fases!altfec
        ' verificar si se trata de un registro completo(alta/baja) o solo de un alta
        If CBool(rs_Fases!estado) Then
            fecbaja = Date  ' solo es un alta ==> tomar el Today (Date)
        Else
            fecbaja = rs_Fases!bajfec   'se trata de un registro completo
        End If
        
        Call Dif_Fechas(fecalta, fecbaja, aux1, aux2, aux3)
        dia = dia + aux1
        Mes = Mes + aux2 + Int(dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        dia = dia Mod 30
        Mes = Mes Mod 12
        
        If Anio = 0 Then
            Call DiasTrab(fecalta, fecbaja, aux1)
            DiasHabiles = DiasHabiles + aux1
        End If
        Antiguedad = 0
End If

If Anio <> 0 Then
    DiasHabiles = 0
End If

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Function


Public Sub bus_interna()
' Busqueda Interna
Dim CampoValor As String 'Campo de la consulta que carga el campo valor

'Dim Param_cur As New ADODB.Recordset
Dim rs_Bus As New ADODB.Recordset
Dim rs_liqvar As New ADODB.Recordset

Dim NuevoValor As String
Dim TipoDeDato As Integer
Dim i As Integer
Dim NombreBuliq As String
Dim NombreCampo As String
Dim Original As String
Dim StringSQL As String
Dim Cargo_Bien As Boolean
Dim pos

    Valor = 0
    Bien = False
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        StringSQL = Trim(Arr_Programa(NroProg).Auxchar)
        ' tengo que eliminar todos los caracteres invalidos del string
        'StrSql = EliminarCHInvalidos(StrSql)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "SQL base del la busqueda interna a ejecutar: " & StringSQL
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & " No se encontr� el nro de busqueda " & NroProg
        End If
        Exit Sub
    End If

On Error GoTo 0 ' desactivo el manejador de errores que exista al momento
On Error GoTo ErrorSQL

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Reemplazo de los parametrso especificados en la sql por sus valores especificos "
End If

'FGZ - 19/08/2004
'pos = InStr(1, UCase(StringSQL), UCase("Buliq_concepto"))
'If pos <> 0 Then
    Call Establecer_Buliq_concepto(Buliq_Concepto(Concepto_Actual).concnro, Cargo_Bien)
    If Not Cargo_Bien Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se pudo cargar Buliq_Concepto para el concepto: " & Buliq_Concepto(Concepto_Actual).concnro
        End If
        Exit Sub
    End If
'End If

'Reemplazo de los parametrso especificados en la consulta por sus valores especificos
StrSql = "SELECT * FROM liqvar"
OpenRecordset StrSql, rs_liqvar

Do Until rs_liqvar.EOF
     Original = "[" & rs_liqvar!lvarnro & "]"
     i = InStr(1, rs_liqvar!lvtablacampo, ".")
     NombreBuliq = Mid(rs_liqvar!lvtablacampo, 1, i - 1)
     NombreCampo = Trim(Mid(rs_liqvar!lvtablacampo, i + 1, Len(rs_liqvar!lvtablacampo)))
     Select Case UCase(NombreBuliq)
     Case "BULIQ_PROCESO":
        TipoDeDato = VarType(buliq_proceso.Fields(NombreCampo))
        Select Case TipoDeDato
            Case 8: 'tipo cadena
                NuevoValor = "'" & buliq_proceso.Fields(NombreCampo) & "'"
            Case 7: 'tipo fecha
                NuevoValor = ConvFecha(buliq_proceso.Fields(NombreCampo))
                'FGZ - 19/03/2004
                If Not CBool(buliq_empleado!empest) Then
                    If buliq_proceso.Fields(NombreCampo) > Empleado_Fecha_Fin Then
                        NuevoValor = ConvFecha(Empleado_Fecha_Fin)
                    End If
                End If
            Case Else: 'cualquier otro tipo
                NuevoValor = buliq_proceso.Fields(NombreCampo)
        End Select
     Case "BULIQ_PERIODO":
        TipoDeDato = VarType(buliq_periodo.Fields(NombreCampo))
        Select Case TipoDeDato
            Case 8: 'tipo cadena
                NuevoValor = "'" & buliq_periodo.Fields(NombreCampo) & "'"
            Case 7: 'tipo fecha
                NuevoValor = ConvFecha(buliq_periodo.Fields(NombreCampo))
                'FGZ - 19/03/2004
                If Not CBool(buliq_empleado!empest) Then
                    If buliq_periodo.Fields(NombreCampo) > Empleado_Fecha_Fin Then
                        NuevoValor = ConvFecha(Empleado_Fecha_Fin)
                    End If
                End If
            Case Else: 'cualquier otro tipo
                NuevoValor = buliq_periodo.Fields(NombreCampo)
        End Select

     Case "BULIQ_EMPLEADO":
        TipoDeDato = VarType(buliq_empleado.Fields(NombreCampo))
        Select Case TipoDeDato
            Case 8: 'tipo cadena
                NuevoValor = "'" & buliq_empleado.Fields(NombreCampo) & "'"
            Case 7: 'tipo fecha
                NuevoValor = ConvFecha(buliq_empleado.Fields(NombreCampo))
            Case Else: 'cualquier otro tipo
                NuevoValor = buliq_empleado.Fields(NombreCampo)
        End Select

     Case "BULIQ_IMPGRALARG":
        TipoDeDato = VarType(buliq_impgralarg.Fields(NombreCampo))
        Select Case TipoDeDato
            Case 8: 'tipo cadena
                NuevoValor = "'" & buliq_impgralarg.Fields(NombreCampo) & "'"
            Case 7: 'tipo fecha
                NuevoValor = ConvFecha(buliq_impgralarg.Fields(NombreCampo))
            Case Else: 'cualquier otro tipo
                NuevoValor = buliq_impgralarg.Fields(NombreCampo)
        End Select

     Case "BULIQ_CABLIQ":
        TipoDeDato = VarType(buliq_cabliq.Fields(NombreCampo))
        Select Case TipoDeDato
            Case 8: 'tipo cadena
                NuevoValor = "'" & buliq_cabliq.Fields(NombreCampo) & "'"
            Case 7: 'tipo fecha
                NuevoValor = ConvFecha(buliq_cabliq.Fields(NombreCampo))
            Case Else: 'cualquier otro tipo
                NuevoValor = buliq_cabliq.Fields(NombreCampo)
        End Select

     Case "BULIQ_CONCEPTO":
        TipoDeDato = VarType(rs_Buliq_Concepto.Fields(NombreCampo))
        Select Case TipoDeDato
            Case 8: 'tipo cadena
                NuevoValor = "'" & rs_Buliq_Concepto.Fields(NombreCampo) & "'"
            Case 7: 'tipo fecha
                NuevoValor = ConvFecha(rs_Buliq_Concepto.Fields(NombreCampo))
            Case Else: 'cualquier otro tipo
                NuevoValor = rs_Buliq_Concepto.Fields(NombreCampo)
        End Select

     End Select
     
     StringSQL = Replace(StringSQL, Original, NuevoValor)

  rs_liqvar.MoveNext
Loop

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Ejecuto el sql que se pas� como parametro y cargo el parametro " & StringSQL
End If

' Ejecuto la consulta que se pas� como parametro y cargo el parametro en
OpenRecordset StringSQL, rs_Bus

If Not rs_Bus.EOF Then
    Bien = True
    'valor = rs_Bus.Fields(CampoValor)
    Valor = CSng(rs_Bus.Fields(0)) ' el orden del campo que quiero
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Retorno el primer campo de la lista del SELECT del sql ejecutado " & Valor
    End If
Else
    Bien = False
End If

Terminar:

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing
If rs_Bus.State = adStateOpen Then rs_Bus.Close
Set rs_Bus = Nothing

Exit Sub

ErrorSQL:
    Bien = False
    If HACE_TRAZA Then
        Call InsertarTraza(NroCab, NroProg, 0, "Busqueda Interna = (" & NroProg & "). Error en el sql: " & Err.Description, 0)
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & " Fall� la busqueda interna "
        Flog.writeline Espacios(Tabulador * 4) & " Error de sintaxis en el sql de la busqueda interna nro: " & NroProg
        Flog.writeline Espacios(Tabulador * 4) & " Error : " & Err.Description
    End If

    GoTo Terminar
End Sub


Public Sub bus_Estructura()
' ---------------------------------------------------------------------------------------------
' Descripcion: Estructura a una Fecha
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoEstr As Long      ' Tipo de Estructura
Dim TipoFecha As Integer    ' 1 - Primer dia del a�o
                            ' 2 - Ultimo dia del a�o
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today
Dim Text As String

'Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Aux_Fecha As Date

   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoEstr = Arr_Programa(NroProg).Auxint1
        TipoFecha = Arr_Programa(NroProg).Auxint2
    Else
        Exit Sub
    End If


Select Case TipoFecha
Case 1: 'Primer dia del a�o
    Aux_Fecha = CDate("01/01/" & Year(Date))
    Texto = "Primer dia del a�o"
Case 2: 'Ultimo dia del a�o
    Aux_Fecha = CDate("31/12/" & Year(Date))
    Texto = "Ultimo dia del a�o"
Case 3: 'Inicio del proceso
    Aux_Fecha = buliq_proceso!profecini
    Texto = "Inicio del proceso"
Case 4: 'Fin del proceso
    Aux_Fecha = buliq_proceso!profecfin
    Texto = "Fin del proceso"
Case 5: 'Inicio del periodo
    Aux_Fecha = buliq_periodo!pliqdesde
    Texto = "Inicio del periodo"
Case 6: 'Fin del periodo
    Aux_Fecha = buliq_periodo!pliqhasta
    Texto = "Fin del periodo"
Case 7: 'Today
    Aux_Fecha = Date
    Texto = "Today"
Case Else
    'tipo de fecha no valido
End Select

If CBool(USA_DEBUG) Then
    Flog.writeline Espacios(Tabulador * 4) & "Busca la estructura  a " & Texto & ": " & Aux_Fecha
End If

'FGZ - 19/03/2004
If Not CBool(buliq_empleado!empest) Then
    If Aux_Fecha > Empleado_Fecha_Fin Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "La fecha de baja del empleado es < a la fecha hasta " & Aux_Fecha & ". Se corre la fecha hasta a: " & Empleado_Fecha_Fin
        End If
        Aux_Fecha = Empleado_Fecha_Fin
    End If
End If

If Not EsNulo(Aux_Fecha) Then
    ' Busco de estructura
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            Valor = rs_Estructura!estrnro
            Bien = True
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro la esructura."
                Flog.writeline Espacios(Tabulador * 4) & "SQL:"
                Flog.writeline Espacios(Tabulador * 5) & StrSql
            End If
        End If
End If
    

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

End Sub


Public Sub bus_AntEstructura()
' ---------------------------------------------------------------------------------------------
' Descripcion: Antiguedad en la Estructura a una Fecha
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoEstr As Long      ' Tipo de Estructura
Dim TipoFecha As Integer    ' 1 - Primer dia del a�o
                            ' 2 - Ultimo dia del a�o
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today
Dim Resultado As Integer    ' Tipo de resultado devuelto
                            ' 1 - En dias
                            ' 2 - En Meses
                            ' 3 - En A�os
                            
'Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Aux_Fecha As Date

Dim FechaDesde As Date
Dim FechaHasta As Date

Dim dia As Integer
Dim Mes As Integer
Dim Anio As Integer
   
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoEstr = Arr_Programa(NroProg).Auxint1
        TipoFecha = Arr_Programa(NroProg).Auxint2
        Resultado = Arr_Programa(NroProg).Auxint3
    Else
        Exit Sub
    End If

Select Case TipoFecha
Case 1: 'Primer dia del a�o
    Aux_Fecha = CDate("01/01/" & Year(Date))
Case 2: 'Ultimo dia del a�o
    Aux_Fecha = CDate("31/12/" & Year(Date))
Case 3: 'Inicio del proceso
    Aux_Fecha = buliq_proceso!profecini
Case 4: 'Fin del proceso
    Aux_Fecha = buliq_proceso!profecfin
Case 5: 'Inicio del periodo
    Aux_Fecha = buliq_periodo!pliqdesde
Case 6: 'Fin del periodo
    Aux_Fecha = buliq_periodo!pliqhasta
Case 7: 'Today
    Aux_Fecha = Date
Case Else
    'tipo de fecha no valido
End Select

'FGZ - 19/03/2004
If Not CBool(buliq_empleado!empest) Then
    If Aux_Fecha > Empleado_Fecha_Fin Then
        Aux_Fecha = Empleado_Fecha_Fin
    End If
End If

If Not EsNulo(Aux_Fecha) Then
    ' Busco de estructura
        StrSql = " SELECT htetdesde,htethasta FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            FechaDesde = rs_Estructura!htetdesde
            FechaHasta = IIf(EsNulo(rs_Estructura!htethasta), Date, rs_Estructura!htethasta)
        End If
        
        Call Dif_Fechas(FechaDesde, Aux_Fecha, aux1, aux2, aux3)
        dia = dia + aux1
        Mes = Mes + aux2 + Int(dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        dia = dia Mod 30
        Mes = Mes Mod 12
        
        Select Case Resultado
        Case 1: ' En dias
            Valor = dia
        Case 2: ' En meses
            Valor = Mes
        Case 3: ' en a�os
            Valor = Anio
        End Select
        Bien = True
End If
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

End Sub

Public Sub bus_ZonaDom()
' ---------------------------------------------------------------------------------------------
' Descripcion: Zona de Domicilio
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoDomi As Long        ' Tipo de Domicilio
Dim TipoEstr As Long        ' Tipo de estructura
Dim Opcion As Long          ' 1 - Empleado
                            ' 2 - Sucursal
                            ' 3 - Empresa
Dim TipoFecha As Integer    ' 1 - Primer dia del a�o
                            ' 2 - Ultimo dia del a�o
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today

'Dim Param_cur As New ADODB.Recordset
Dim rs_Zona As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Aux_Fecha As Date

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoDomi = Arr_Programa(NroProg).Auxint1
        Opcion = Arr_Programa(NroProg).Auxint2
        Select Case Opcion
        Case 1:
            TipoEstr = 1
        Case 2:
            TipoEstr = 1
            TipoFecha = Arr_Programa(NroProg).Auxint3
        Case 3:
            TipoEstr = 10
            TipoFecha = Arr_Programa(NroProg).Auxint3
        Case Else
        End Select
    Else
        Exit Sub
    End If

    
    Select Case TipoFecha
    Case 1: 'Primer dia del a�o
        Aux_Fecha = CDate("01/01/" & Year(Date))
    Case 2: 'Ultimo dia del a�o
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 3: 'Inicio del proceso
        Aux_Fecha = buliq_proceso!profecini
    Case 4: 'Fin del proceso
        Aux_Fecha = buliq_proceso!profecfin
    Case 5: 'Inicio del periodo
        Aux_Fecha = buliq_periodo!pliqdesde
    Case 6: 'Fin del periodo
        Aux_Fecha = buliq_periodo!pliqhasta
    Case 7: 'Today
        Aux_Fecha = Date
    Case Else
        'tipo de fecha no valido
    End Select

    'FGZ - 19/03/2004
    If Not CBool(buliq_empleado!empest) Then
        If Aux_Fecha > Empleado_Fecha_Fin Then
            Aux_Fecha = Empleado_Fecha_Fin
        End If
    End If
    
    ' De acuerdo a la opcion busco la zona
    Select Case Opcion
    Case 1: 'Empleado
        StrSql = " SELECT zonanro FROM detdom " & _
                 " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                 " WHERE cabdom.ternro = " & buliq_empleado!ternro & " AND " & _
                 " cabdom.domdefault = -1"
        OpenRecordset StrSql, rs_Zona
        If Not rs_Zona.EOF Then
            Valor = rs_Zona!zonanro
            Bien = True
        End If

    Case 2: 'Sucursal
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            StrSql = " SELECT ternro FROM sucursal " & _
                     " WHERE estrnro =" & rs_Estructura!estrnro
            OpenRecordset StrSql, rs_Sucursal
            
            If Not rs_Sucursal.EOF Then
                StrSql = " SELECT zonanro FROM detdom " & _
                         " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                         " WHERE cabdom.ternro = " & rs_Sucursal!ternro & " AND " & _
                         " cabdom.domdefault = -1 "
                OpenRecordset StrSql, rs_Zona
                If Not rs_Zona.EOF Then
                    Valor = rs_Zona!zonanro
                    Bien = True
                End If
            End If  ' If Not rs_Sucursal.EOF Then
        End If ' If Not rs_Estructura.EOF Then
    
    Case 3: 'Empresa
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            StrSql = " SELECT ternro FROM empresa " & _
                     " WHERE estrnro =" & rs_Estructura!estrnro
            OpenRecordset StrSql, rs_Sucursal
            
            If Not rs_Sucursal.EOF Then
                StrSql = " SELECT zonanro FROM detdom " & _
                         " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                         " WHERE cabdom.ternro = " & rs_Sucursal!ternro & " AND " & _
                         " cabdom.domdefault = -1 "
                OpenRecordset StrSql, rs_Zona
                If Not rs_Zona.EOF Then
                    Valor = rs_Zona!zonanro
                    Bien = True
                End If
            End If  ' If Not rs_Sucursal.EOF Then
        End If ' If Not rs_Estructura.EOF Then
    
    Case Else
    End Select
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Zona.State = adStateOpen Then rs_Zona.Close
Set rs_Zona = Nothing

If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
Set rs_Sucursal = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

End Sub


Public Sub bus_ZonaDom_old()
' ---------------------------------------------------------------------------------------------
' Descripcion: Zona de Domicilio
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoDomi As Long        ' Tipo de Domicilio
Dim TipoEstr As Long        ' Tipo de estructura
Dim Opcion As Long          ' 1 - Empleado
                            ' 2 - Sucursal
                            ' 3 - Empresa
Dim TipoFecha As Integer    ' 1 - Primer dia del a�o
                            ' 2 - Ultimo dia del a�o
                            ' 3 - Inicio del proceso
                            ' 4 - Fin del proceso
                            ' 5 - Inicio del Periodo
                            ' 6 - Fin del periodo
                            ' 7 - Today

'Dim Param_cur As New ADODB.Recordset
Dim rs_Zona As New ADODB.Recordset
Dim rs_Sucursal As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim Aux_Fecha As Date

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoDomi = Arr_Programa(NroProg).Auxint1
        Opcion = Arr_Programa(NroProg).Auxint2
        Select Case Opcion
        Case 1:
            TipoEstr = 1
        Case 2:
            TipoEstr = 1
            TipoFecha = Arr_Programa(NroProg).Auxint3
        Case 3:
            TipoEstr = 10
            TipoFecha = Arr_Programa(NroProg).Auxint3
        Case Else
        End Select
    Else
        Exit Sub
    End If

    
    Select Case TipoFecha
    Case 1: 'Primer dia del a�o
        Aux_Fecha = CDate("01/01/" & Year(Date))
    Case 2: 'Ultimo dia del a�o
        Aux_Fecha = CDate("31/12/" & Year(Date))
    Case 3: 'Inicio del proceso
        Aux_Fecha = buliq_proceso!profecini
    Case 4: 'Fin del proceso
        Aux_Fecha = buliq_proceso!profecfin
    Case 5: 'Inicio del periodo
        Aux_Fecha = buliq_periodo!pliqdesde
    Case 6: 'Fin del periodo
        Aux_Fecha = buliq_periodo!pliqhasta
    Case 7: 'Today
        Aux_Fecha = Date
    Case Else
        'tipo de fecha no valido
    End Select

    'FGZ - 19/03/2004
    If Not CBool(buliq_empleado!empest) Then
        If Aux_Fecha > Empleado_Fecha_Fin Then
            Aux_Fecha = Empleado_Fecha_Fin
        End If
    End If
    
    ' De acuerdo a la opcion busco la zona
    Select Case Opcion
    Case 1: 'Empleado
        StrSql = " SELECT zonanro FROM detdom " & _
                 " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                 " WHERE cabdom.ternro = " & buliq_empleado!ternro & " AND " & _
                 " cabdom.tipnro =" & TipoEstr
        OpenRecordset StrSql, rs_Zona
        If Not rs_Zona.EOF Then
            Valor = rs_Zona!Fields(0)
            Bien = True
        End If

    Case 2, 3: 'Sucursal o Empresa
        StrSql = " SELECT estrnro FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & TipoEstr & " AND " & _
                 " (htetdesde <= " & ConvFecha(Aux_Fecha) & ") AND " & _
                 " ((" & ConvFecha(Aux_Fecha) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
    
        If Not rs_Estructura.EOF Then
            StrSql = " SELECT ternro FROM sucursal " & _
                     " WHERE estrnro =" & rs_Estructura!estrnro
            OpenRecordset StrSql, rs_Sucursal
            
            If Not rs_Sucursal.EOF Then
                StrSql = " SELECT zonanro FROM detdom " & _
                         " INNER JOIN cabdom ON detdom.domnro = cabdom.domnro " & _
                         " WHERE cabdom.ternro = " & rs_Sucursal!ternro & " AND " & _
                         " cabdom.tipnro =" & TipoEstr
                OpenRecordset StrSql, rs_Zona
                If Not rs_Zona.EOF Then
                    Valor = rs_Zona!zonanro
                    Bien = True
                End If
            End If  ' If Not rs_Sucursal.EOF Then
        End If ' If Not rs_Estructura.EOF Then
    Case Else
    End Select
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Zona.State = adStateOpen Then rs_Zona.Close
Set rs_Zona = Nothing

If rs_Sucursal.State = adStateOpen Then rs_Sucursal.Close
Set rs_Sucursal = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

End Sub


Public Sub bus_PagoDtoLic()
' ---------------------------------------------------------------------------------------------
' Descripcion: Pago / Descuento de Licencias
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Opcion As Integer       ' 1 - Pago de adelanto
                            ' 2 - Descuento de adelanto
                            ' 3 - Descuento de vacaciones
                            ' 4 - Descuento de vacaciones
Dim Todos As Boolean        'TRUE - Todas los tipos de licencias y
                            'FALSE - solo el tipo especificado en TipoLicencia
Dim TipoLicencia As Long    'tipdia.tdnro

'Dim Param_cur As New ADODB.Recordset
Dim rs_Vacpagdesc As New ADODB.Recordset
Dim Texto As String
   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
   
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Opcion = Arr_Programa(NroProg).Auxint1
        Todos = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not Todos Then
            TipoLicencia = Arr_Programa(NroProg).Auxint2
        End If
        If CBool(USA_DEBUG) Then
            Select Case Opcion
                Case 1:
                    Texto = "Pago de adelanto"
                Case 2:
                    Texto = "Descuento de adelanto"
                Case 3:
                    Texto = "Pago de vacaciones"
                Case 4:
                    Texto = "Descuento de vacaciones"
            End Select
            Flog.writeline Espacios(Tabulador * 5) & "Busco " & Texto
            If Todos Then
                Flog.writeline Espacios(Tabulador * 5) & "para todos los tipos de licencias"
            Else
                Flog.writeline Espacios(Tabulador * 5) & "para el tipos de licencia " & TipoLicencia
            End If
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

    If Todos Then
        StrSql = " SELECT vacpdnro,cantdias FROM vacpagdesc " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " pliqnro =" & buliq_proceso!PliqNro & " AND " & _
                 " tprocnro =" & buliq_proceso!tprocnro & " AND " & _
                 " pago_dto =" & Opcion
    Else
        StrSql = " SELECT vacpdnro,cantdias FROM vacpagdesc " & _
                 " INNER JOIN emp_lic ON vacpagdesc.emp_licnro = emp_lic.emp_licnro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " pliqnro =" & buliq_proceso!PliqNro & " AND " & _
                 " tprocnro =" & buliq_proceso!tprocnro & " AND " & _
                 " emp_lic.tdnro =" & TipoLicencia & " AND " & _
                 " pago_dto =" & Opcion
    End If
    OpenRecordset StrSql, rs_Vacpagdesc

    If Not rs_Vacpagdesc.EOF Then
        Bien = True
        Valor = rs_Vacpagdesc!cantdias
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Se encontr� Pagos/Dtos por " & Valor & " dias "
            Flog.writeline Espacios(Tabulador * 4) & "Marco con el pronro del proceso actual"
        End If
        
        'Marco con el pronro del proceso actual
        StrSql = "UPDATE vacpagdesc SET pronro = " & buliq_proceso!pronro & _
                 " WHERE vacpdnro = " & rs_Vacpagdesc!vacpdnro
        objConn.Execute StrSql, , adExecuteNoRecords
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontraron Pagos/Dtos "
        End If
    End If
        
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Vacpagdesc.State = adStateOpen Then rs_Vacpagdesc.Close
Set rs_Vacpagdesc = Nothing

End Sub

Public Sub bus_PagoDtoLic_OLD()
' ---------------------------------------------------------------------------------------------
' Descripcion: Pago / Descuento de Licencias
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Opcion As Integer       ' 1 - Pago de adelanto
                            ' 2 - Descuento de adelanto
                            ' 3 - Descuento de vacaciones
                            ' 4 - Descuento de vacaciones
Dim Todos As Boolean        'TRUE - Todas los tipos de licencias y
                            'FALSE - solo el tipo especificado en TipoLicencia
Dim TipoLicencia As Long    'tipdia.tdnro

'Dim Param_cur As New ADODB.Recordset
Dim rs_Vacpagdesc As New ADODB.Recordset

   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Opcion = Arr_Programa(NroProg).Auxint1
        Todos = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not Todos Then
            TipoLicencia = Arr_Programa(NroProg).Auxint2
        End If
    Else
        Exit Sub
    End If

    If Todos Then
        StrSql = " SELECT cantdias FROM vacpagdesc " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " pliqnro =" & buliq_proceso!PliqNro & " AND " & _
                 " tprocnro =" & buliq_proceso!tprocnro & " AND " & _
                 " concnro =" & Buliq_Concepto(Concepto_Actual).concnro & " AND " & _
                 " pago_dto =" & Opcion
    Else
        StrSql = " SELECT cantdias FROM vacpagdesc " & _
                 " INNER JOIN emp_lic ON vacpagdesc.emp_licnro = emp_lic.emp_licnro " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " pliqnro =" & buliq_proceso!PliqNro & " AND " & _
                 " tprocnro =" & buliq_proceso!tprocnro & " AND " & _
                 " concnro =" & Buliq_Concepto(Concepto_Actual).concnro & " AND " & _
                 " emp_lic.tdnro =" & TipoLicencia & " AND " & _
                 " pago_dto =" & Opcion
    End If
    OpenRecordset StrSql, rs_Vacpagdesc

    If Not rs_Vacpagdesc.EOF Then
        Bien = True
        Valor = rs_Vacpagdesc!cantdias
    End If
        
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Vacpagdesc.State = adStateOpen Then rs_Vacpagdesc.Close
Set rs_Vacpagdesc = Nothing

End Sub


Public Sub bus_NovGTI()
' ---------------------------------------------------------------------------------------------
' Descripcion: Novedades de GTI
' Autor      : FGZ
' Fecha      : 25/11/2003
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim NroConc As Long             ' concepto.concnro

Dim NroParametro As Long        ' tipopar.tpanro

Dim Opcion As Integer           ' 1 - Aprobadas
                                ' 2 - No Aprobadas
                                ' 3 - Todas
                                 
Dim DesdeHasta As Integer       ' 1- Proceso
                                ' 2- Periodo
                                
Dim ModeloGTI As Boolean        ' 0 - Uno
                                '-1 - Todos
                                
Dim Modelo As Long              'si ModeloGTI = 0 ==> gti_tipproc

'Dim Param_cur As New ADODB.Recordset
Dim rs_AcuNov As New ADODB.Recordset
Dim rs_gti_procacum As New ADODB.Recordset
 
Dim FechaDesde As Date
Dim FechaHasta As Date

Dim Aux_Suma As Single

    Bien = False
    Valor = 0
   
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroConc = Arr_Programa(NroProg).Auxint1
        NroParametro = Arr_Programa(NroProg).Auxint2
        Opcion = Arr_Programa(NroProg).Auxint3
        DesdeHasta = Arr_Programa(NroProg).Auxint4
        ModeloGTI = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not ModeloGTI Then
            Modelo = Arr_Programa(NroProg).Auxlog5
        End If
    Else
        Exit Sub
    End If

    If DesdeHasta = 1 Then
        FechaDesde = buliq_proceso!profecini
        FechaHasta = buliq_proceso!profecfin
    Else
        FechaDesde = buliq_periodo!pliqdesde
        FechaHasta = buliq_periodo!pliqhasta
    End If

    If Not ModeloGTI Then
        StrSql = " SELECT gpanro, gpadesde, gpahasta " & _
                 " FROM gti_Procacum " & _
                 " WHERE  gpadesde >= " & FechaDesde & " AND gpahasta <= " & FechaHasta
    Else
        StrSql = " SELECT gpanro, gpadesde, gpahasta " & _
                 " FROM gti_Procacum " & _
                 " WHERE  gpadesde >= " & FechaDesde & " AND gpahasta <= " & FechaHasta & _
                 " AND gtprocnro = " & Modelo
    End If
    OpenRecordset StrSql, rs_gti_procacum

    Aux_Suma = 0
    Do While Not rs_gti_procacum.EOF
               StrSql = "SELECT * FROM gti_acunov " & _
                     " WHERE gti_acunov.concnro = " & NroConc & _
                     " AND gti_acunov.tpanro = " & NroParametro & _
                     " AND gti_acunov.tternro = " & buliq_empleado!ternro & _
                     " AND not EsNulo( gti_acunov.acnovfecaprob)" & _
                     " AND gti_acunov.gpanro = " & rs_gti_procacum!gpanro
                OpenRecordset StrSql, rs_AcuNov
                
                Do While Not rs_AcuNov.EOF
                
                    Aux_Suma = Aux_Suma + rs_AcuNov!acnovvalor
                    
                    rs_AcuNov.MoveNext
                Loop
                     
        rs_gti_procacum.MoveNext
    Loop

    Bien = False
    Valor = Aux_Suma


' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_AcuNov.State = adStateOpen Then rs_AcuNov.Close
Set rs_AcuNov = Nothing

If rs_gti_procacum.State = adStateOpen Then rs_gti_procacum.Close
Set rs_gti_procacum = Nothing

End Sub

Public Sub bus_DiasHabiles_ConFases()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias Habiles entre dos fechas teniendo en cuenta las fases
' Autor      : FGZ
' Fecha      : 05/01/2004
' Ultima Mod.: FGZ - 27/01/2005
' Descripcion: se le agrego que tenga en cuenta las fases
' ---------------------------------------------------------------------------------------------
Dim TipoFecha As Integer    '1- Periodo
                            '2- Proceso
Dim DiasHabiles As Single
Dim dia As Date
Dim EsFeriado As Boolean

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim objFeriado As New Feriado
Dim ConFases As Boolean
Dim Aux_FechaDesde As Date
Dim Aux_FechaHasta As Date
Dim rs_Fases As New ADODB.Recordset

Dim IncluyeSabados As Boolean
Dim PorcentajeSabados As Single
Dim IncluyeFeriados As Boolean
Dim PorcentajeFeriados As Single

'inicializacion de variables
    ConFases = True
    DiasHabiles = 0
    Bien = False
    
    Set objFeriado.Conexion = objConn
    
    ' Obtener los parametros de la Busqueda
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoFecha = Arr_Programa(NroProg).Auxint1
        ConFases = CBool(Arr_Programa(NroProg).Auxlog1)
        IncluyeSabados = CBool(Arr_Programa(NroProg).Auxlog2)
        PorcentajeSabados = Arr_Programa(NroProg).Auxint2
        IncluyeFeriados = CBool(Arr_Programa(NroProg).Auxlog3)
        PorcentajeFeriados = Arr_Programa(NroProg).Auxint3
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada "
        End If
        Exit Sub
    End If

    If TipoFecha = 1 Then
        FechaDesde = buliq_periodo!pliqdesde
        FechaHasta = buliq_periodo!pliqhasta
    Else
        FechaDesde = buliq_proceso!profecini
        FechaHasta = buliq_proceso!profecfin
    End If
    
If ConFases Then
    'Busco las fases del periodo
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND real = -1 AND Fases.altfec <= " & ConvFecha(FechaHasta) & _
             " AND (Fases.bajfec >= " & ConvFecha(FechaDesde) & " OR Fases.bajfec is null )" & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    Do While Not rs_Fases.EOF
        Aux_FechaDesde = IIf(rs_Fases!altfec < FechaDesde, FechaDesde, rs_Fases!altfec)
        If Not EsNulo(rs_Fases!bajfec) Then
            Aux_FechaHasta = IIf(rs_Fases!bajfec < FechaHasta, rs_Fases!bajfec, FechaHasta)
        Else
            Aux_FechaHasta = FechaHasta
        End If
        
        dia = Aux_FechaDesde
        Do While dia <= Aux_FechaHasta
            
            EsFeriado = objFeriado.Feriado(dia, buliq_empleado!ternro, False)
            
            If Not EsFeriado Then   'No es feriado
                If Not Weekday(dia) = 1 Then 'Domingo
                    If Weekday(dia) = 7 Then 'Sabado
                        If IncluyeSabados Then
                            DiasHabiles = DiasHabiles + (1 * PorcentajeSabados / 100)
                        End If
                    Else
                        DiasHabiles = DiasHabiles + 1
                    End If
                End If
            Else    'Incluye feriados
                DiasHabiles = DiasHabiles + (1 * PorcentajeFeriados / 100)
            End If
            
            dia = dia + 1
        Loop
        
       rs_Fases.MoveNext
    Loop
Else
    dia = FechaDesde
    Do While dia <= FechaHasta
        
        EsFeriado = objFeriado.Feriado(dia, buliq_empleado!ternro, False)
        If Not EsFeriado Then   'No es feriado
            If Not Weekday(dia) = 1 Then 'Domingo
                If Weekday(dia) = 7 Then 'Sabado
                    If IncluyeSabados Then
                        DiasHabiles = DiasHabiles + (1 * PorcentajeSabados / 100)
                    End If
                Else
                    DiasHabiles = DiasHabiles + 1
                End If
            End If
        Else    'Incluye feriados
            DiasHabiles = DiasHabiles + (1 * PorcentajeFeriados / 100)
        End If
        
        
'        If Not EsFeriado And Not Weekday(dia) = 1 Then
'            ' No es feriado no Domingo
'            If Weekday(dia) = 7 Then 'Sabado
'                DiasHabiles = DiasHabiles + 0.5
'            Else
'                DiasHabiles = DiasHabiles + 1
'            End If
'        End If
        dia = dia + 1
    Loop
End If

Bien = True
Valor = DiasHabiles

'Cierro y libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub




Public Sub bus_DiasHabiles()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias Habiles entre dos fechas
' Autor      : FGZ
' Fecha      : 05/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TipoFecha As Integer    '1- Periodo
                            '2- Proceso
Dim DiasHabiles As Single
Dim dia As Date
Dim EsFeriado As Boolean

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim objFeriado As New Feriado

'Dim Param_cur As New ADODB.Recordset

   
' inicializacion de variables
    Set objFeriado.Conexion = objConn
    'Set objFeriado.ConexionTraza = objConn

    Bien = False
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoFecha = Arr_Programa(NroProg).Auxint1
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada "
        End If
        Exit Sub
    End If

    If TipoFecha = 1 Then
        FechaDesde = buliq_periodo!pliqdesde
        FechaHasta = buliq_periodo!pliqhasta
    Else
        FechaDesde = buliq_proceso!profecini
        FechaHasta = buliq_proceso!profecfin
    End If
    
dia = FechaDesde
Do While dia <= FechaHasta
    
    EsFeriado = objFeriado.Feriado(dia, buliq_empleado!ternro, HACE_TRAZA)
    
    If Not EsFeriado And Not Weekday(dia) = 1 Then
        ' No es feriado no Domingo
        If Weekday(dia) = 7 Then 'Sabado
            DiasHabiles = DiasHabiles + 0.5
        Else
            DiasHabiles = DiasHabiles + 1
        End If
    End If
    dia = dia + 1
Loop
Bien = True
Valor = DiasHabiles
End Sub



'Public Sub bus_GTI_DiasTrabajados()
'' ---------------------------------------------------------------------------------------------
'' Descripcion: Sumariza los dias trabajados (GTI)
'' segun empaque, productoy embalador y tercero
'' Autor      : FGZ
'' Fecha      : 05/01/2004
'' Ultima Mod.:
'' Descripcion:
'' ---------------------------------------------------------------------------------------------
'Dim NroProducto As Long                 ' de gti
'Dim NroEmpaque As Long                  ' de gti
'Dim HoraProduccion As Long              ' de gti
'Dim CategoriasEmbaladores As String     'GTI (210,211,212)
'
''Dim Param_cur As New ADODB.Recordset
'Dim rs_AcuNov As New ADODB.Recordset
'Dim rs_gti_achdiario As New ADODB.Recordset
'Dim rs_gti_achMensual As New ADODB.Recordset
'Dim rs_Estructura As New ADODB.Recordset
'
'Dim FechaDesde As Date
'Dim FechaHasta As Date
'
'Dim dias_trabajados As Integer
'
'    Bien = False
'    valor = 0
'    CategoriasEmbaladores = "210,211,212"
'
'    ' Obtener los parametros de la Busqueda
'    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        NroProducto = Arr_Programa(nroprog).auxint1
'        NroEmpaque = Arr_Programa(nroprog).auxint2
'        HoraProduccion = Arr_Programa(nroprog).auxint3
'        CategoriasEmbaladores = Arr_Programa(nroprog).auxchar1
'    Else
'        Exit Sub
'    End If
'
'    FechaDesde = buliq_periodo!pliqdesde
'    FechaHasta = buliq_periodo!pliqhasta
'
'    StrSql = " SELECT estrnro FROM his_estructura " & _
'             " WHERE ternro = " & buliq_empleado!Ternro & " AND " & _
'             " tenro = " & NroEmpaque & " AND " & _
'             " (htetdesde <= " & ConvFecha(FechaDesde) & ") AND " & _
'             " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))"
'    OpenRecordset StrSql, rs_Estructura
'
'    If rs_Estructura.EOF Then
'        'Flog "No se encuentra la sucursal"
'        Exit Sub
'    End If
'
'    'Recorriendo el desglose diario
'    StrSql = "SELECT * FROM gti_achdiario " & _
'         " WHERE (" & ConvFecha(FechaDesde) & " <= gti_achdiario.achdfecha) " & _
'         " AND (gti_achdiario.achdfecha <= " & ConvFecha(FechaHasta) & ")" & _
'         " AND (gti_achdiario.puenro = " & NroProducto & ") " & _
'         " AND (gti_achdiario.sucursal = " & rs_Estructura!estrnro & ")" & _
'         " AND (ternro = " & buliq_empleado!Ternro & ") AND (gti_achdiario.thnro =" & HoraProduccion & " )"
'    OpenRecordset StrSql, rs_gti_achdiario
'
'    Do While Not rs_gti_achdiario.EOF
'
'        dias_trabajados = dias_trabajados + rs_gti_achdiario!achdcanthoras
'
'        rs_gti_achdiario.MoveNext
'    Loop
'
'    'Recorriendo el desglose acumulado, teoricamente deberia ser un registro
'    ' Pendiente para mejorar performance
''    StrSql = "SELECT * FROM gti_achmensual " & _
''         " WHERE achmano = " & buliq_periodo!pliqnro & _
''         "AND catnro in (" & CategoriasEmbaladores & ")" & _
''         " AND (gti_achdiario.puenro = " & NroProducto & ") " & _
''         " AND (gti_achdiario.sucursal = " & rs_Estructura.estrnro & ")" & _
''         " AND (ternro = " & buliq_empleado!ternro & ")" & _
''         " AND (gti_achdiario.thnro =" & HoraProduccion & " )"
''    OpenRecordset StrSql, rs_gti_achMensual
''
''    Do While Not rs_gti_achMensual.EOF
''
''        dias_trabajados = dias_trabajados + rs_gti_achMensual!achmcanthoras
''
''        rs_gti_achMensual.MoveNext
''    Loop
'
'    valor = dias_trabajados
'    Bien = True
'
'' Cierro todo y libero
''If Param_cur.State = adStateOpen Then Param_cur.Close
''Set Param_cur = Nothing
'
'If rs_AcuNov.State = adStateOpen Then rs_AcuNov.Close
'Set rs_AcuNov = Nothing
'
'If rs_gti_achdiario.State = adStateOpen Then rs_gti_achdiario.Close
'Set rs_gti_achdiario = Nothing
'
'If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
'Set rs_Estructura = Nothing
'
'End Sub




Public Sub bus_Licencias()
' ---------------------------------------------------------------------------------------------
' Descripcion: Dias de Licencias entre dos fechas (de un tipo o de todos los tipos)
' Autor      : FGZ
' Fecha      : 14/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoLicencia As Long    'Tipo de Estructura
Dim Todas As Boolean        'Todos los tipos

Dim dias As Integer
Dim SumaDias As Integer
Dim SumaDiasYaGenerados As Integer
Dim FechaDeInicio As Date
Dim FechaDeFin As Date
Dim TipoDia_Ok As Boolean
Dim Dias_Mes_Anterior As Integer

Dim rs_Estructura As New ADODB.Recordset
Dim rs_tipd_con As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset


    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Todas = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not Todas Then
            TipoLicencia = Arr_Programa(NroProg).Auxint1
        End If
    Else
        Exit Sub
    End If

'FGZ - 29/01/2004
FechaDeInicio = buliq_proceso!profecini
FechaDeFin = buliq_proceso!profecfin

' Primero Busco  los tipos de dias asociados a los conceptos
If Todas Then 'Todos los tipos de Licencias
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro
Else 'Todos las Licencias del tipo especificado
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro & _
             " AND tdnro = " & TipoLicencia
End If
OpenRecordset StrSql, rs_tipd_con

Do While Not rs_tipd_con.EOF
    TipoDia_Ok = True
    If Not EsNulo(rs_tipd_con!tenro) Then
        If rs_tipd_con!tenro <> 0 Then
            StrSql = " SELECT * FROM his_estructura " & _
                     " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                     " tenro =" & rs_tipd_con!tenro & " AND " & _
                     " estrnro = " & rs_tipd_con!estrnro & " AND " & _
                     " (htetdesde <= " & ConvFecha(FechaDeFin) & ") AND " & _
                     " ((" & ConvFecha(FechaDeFin) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If rs_Estructura.EOF Then
                TipoDia_Ok = False
            End If
        End If
    End If

    If CBool(TipoDia_Ok) Then
        StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
                 " AND tdnro =" & rs_tipd_con!tdnro & _
                 " AND elfechadesde <=" & ConvFecha(FechaDeFin) & _
                 " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
        OpenRecordset StrSql, rs_Lic
        
        dias = 0
        Do While Not rs_Lic.EOF
            dias = CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
            'reviso si la licencia es completa
            If Todas Then 'Todos los tipos de Licencias
                Dias_Mes_Anterior = Dias_Licencias_Mes_Anterior(buliq_empleado!ternro, DateAdd("m", -1, FechaDeInicio), FechaDeInicio - 1)
                If Dias_Mes_Anterior = 30 Then
                    'calculo los dias reales del mes
                    Dias_Mes_Anterior = DateDiff("d", DateAdd("m", -1, FechaDeInicio), FechaDeInicio - 1) + 1
                    dias = dias + (Dias_Mes_Anterior - 30)
                End If
            Else
                ' solo este tipo
                If rs_Lic!elfechadesde <= DateAdd("m", -1, FechaDeInicio) Then
                    If rs_Lic!elfechahasta >= DateAdd("m", -1, FechaDeFin) Then
                        'Para ajustar la cantidad de dias cuando la lic sobrepasa al mes y fue topeada
                        Dias_Mes_Anterior = DateDiff("d", DateAdd("m", -1, FechaDeInicio), FechaDeInicio - 1) + 1
                        dias = dias + (Dias_Mes_Anterior - 30)
                    End If
                End If
            End If
            SumaDias = SumaDias + dias
            
            'Marco la licencia para que no se pueda Borrar
            StrSql = "UPDATE emp_lic SET pronro = " & NroProc & _
                     " WHERE emp_licnro = " & rs_Lic!emp_licnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_Lic.MoveNext
        Loop
    End If
    rs_tipd_con.MoveNext
Loop

' --------------------------------------------
' FGZ - 29/01/2004
' Buscar todas las licencias
' Busco los detliq (campo dlicant "cantidad") del periodo cuyas licencias emp_lic esten marcadas (en pronro)
' este valor +  SumaDias no debe seperar 30 dias
' FGZ - 29/01/2004
' QUEDA PENDIENTE
' --------------------------------------------

If Month(FechaDeInicio) = 2 Then 'Febrero
    If Biciesto(Year(FechaDeInicio)) Then
        If SumaDias >= 29 Then
            Valor = 30
        Else
            Valor = SumaDias
        End If
    Else
        If SumaDias >= 28 Then
            Valor = 30
        Else
            Valor = SumaDias
        End If
    End If
Else
    If SumaDias > 30 Then
        Valor = 30
    Else
        Valor = SumaDias
    End If
End If
Bien = True
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

If rs_Lic.State = adStateOpen Then rs_Lic.Close
Set rs_Lic = Nothing

If rs_tipd_con.State = adStateOpen Then rs_tipd_con.Close
Set rs_tipd_con = Nothing

End Sub



Public Sub bus_LicenciasMesCalendario()
' ---------------------------------------------------------------------------------------------
' Descripcion: Dias de Licencias entre dos fechas (de un tipo o de todos los tipos), sin topeo de dias
' Autor      : FGZ
' Fecha      : 12/03/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoLicencia As Long    'Tipo de Estructura
Dim Todas As Boolean        'Todos los tipos

Dim dias As Integer
Dim SumaDias As Integer
Dim SumaDiasYaGenerados As Integer
Dim FechaDeInicio As Date
Dim FechaDeFin As Date
Dim TipoDia_Ok As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_tipd_con As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset


    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Todas = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not Todas Then
            TipoLicencia = Arr_Programa(NroProg).Auxint1
        End If
    Else
        Exit Sub
    End If


'FechaDeInicio = buliq_periodo!pliqdesde
'FechaDeFin = buliq_periodo!pliqhasta

'FGZ - 29/01/2004
FechaDeInicio = buliq_proceso!profecini
FechaDeFin = buliq_proceso!profecfin

' Primero Busco  los tipos de dias asociados a los conceptos
If Todas Then 'Todos los tipos de Licencias
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro
Else 'Todos las Licencias del tipo especificado
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro & _
             " AND tdnro = " & TipoLicencia
End If
OpenRecordset StrSql, rs_tipd_con

Do While Not rs_tipd_con.EOF
    TipoDia_Ok = True
    If Not EsNulo(rs_tipd_con!tenro) Then
        StrSql = " SELECT * FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & rs_tipd_con!tenro & " AND " & _
                 " estrnro = " & rs_tipd_con!estrnro & " AND " & _
                 " (htetdesde <= " & ConvFecha(FechaDeFin) & ") AND " & _
                 " ((" & ConvFecha(FechaDeFin) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        
        If rs_Estructura.EOF Then
            TipoDia_Ok = False
        End If
    End If

    If CBool(TipoDia_Ok) Then
'        StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
'                 " AND tdnro =" & rs_tipd_con!tdnro & _
'                 " AND elfechadesde >=" & ConvFecha(FechaDeInicio) & _
'                 " AND elfechahasta <= " & ConvFecha(FechaDeFin)
        
        StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
                 " AND tdnro =" & rs_tipd_con!tdnro & _
                 " AND elfechadesde <=" & ConvFecha(FechaDeFin) & _
                 " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
        OpenRecordset StrSql, rs_Lic
        
        dias = 0
        Do While Not rs_Lic.EOF
            dias = CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
            SumaDias = SumaDias + dias
            
            'Marco la licencia para que no se pueda Borrar
            StrSql = "UPDATE emp_lic SET pronro = " & NroProc & _
                     " WHERE emp_licnro = " & rs_Lic!emp_licnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_Lic.MoveNext
        Loop
    End If
    rs_tipd_con.MoveNext
Loop

' --------------------------------------------
' FGZ - 29/01/2004
' Buscar todas las licencias
' Busco los detliq (campo dlicant "cantidad") del periodo cuyas licencias emp_lic esten marcadas (en pronro)
' este valor +  SumaDias no debe seperar 30 dias
' FGZ - 29/01/2004
' QUEDA PENDIENTE
' --------------------------------------------
'If Month(FechaDeInicio) = 2 Then 'Febrero
'    If Biciesto(Year(FechaDeInicio)) Then
'        If SumaDias >= 29 Then
'            Valor = 30
'        Else
'            Valor = SumaDias
'        End If
'    Else
'        If SumaDias >= 28 Then
'            Valor = 30
'        Else
'            Valor = SumaDias
'        End If
'    End If
'Else
'    If SumaDias > 30 Then
'        Valor = 30
'    Else
'        Valor = SumaDias
'    End If
'End If

Valor = SumaDias
Bien = True
    

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

If rs_Lic.State = adStateOpen Then rs_Lic.Close
Set rs_Lic = Nothing

If rs_tipd_con.State = adStateOpen Then rs_tipd_con.Close
Set rs_tipd_con = Nothing

End Sub


Public Sub bus_LicenciasPeriodoGTI()
' ---------------------------------------------------------------------------------------------
' Descripcion: Dias de Licencias entre dos fechas (de un tipo o de todos los tipos), sin topeo de dias
'               segun periodo de GTI
' Autor      : FGZ
' Fecha      : 09/06/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoLicencia As Long    'Tipo de Estructura
Dim Todas As Boolean        'Todos los tipos

Dim dias As Integer
Dim SumaDias As Integer
Dim SumaDiasYaGenerados As Integer
Dim FechaDeInicio As Date
Dim FechaDeFin As Date
Dim TipoDia_Ok As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_tipd_con As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset
Dim rs_PeriodoGTI As New ADODB.Recordset

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Todas = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not Todas Then
            TipoLicencia = Arr_Programa(NroProg).Auxint1
        End If
    Else
        Exit Sub
    End If

' Busco el periodo de GTI
StrSql = "SELECT * FROM gti_per "
StrSql = StrSql & " WHERE pgtimes = " & buliq_periodo!pliqmes
StrSql = StrSql & " AND pgtianio = " & buliq_periodo!pliqanio
OpenRecordset StrSql, rs_PeriodoGTI

If Not rs_PeriodoGTI.EOF Then
    FechaDeInicio = rs_PeriodoGTI!pgtidesde
    FechaDeFin = rs_PeriodoGTI!pgtihasta
End If


' Primero Busco  los tipos de dias asociados a los conceptos
If Todas Then 'Todos los tipos de Licencias
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro
Else 'Todos las Licencias del tipo especificado
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro & _
             " AND tdnro = " & TipoLicencia
End If
OpenRecordset StrSql, rs_tipd_con

Do While Not rs_tipd_con.EOF
    TipoDia_Ok = True
    If Not EsNulo(rs_tipd_con!tenro) Then
        StrSql = " SELECT * FROM his_estructura " & _
                 " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                 " tenro =" & rs_tipd_con!tenro & " AND " & _
                 " estrnro = " & rs_tipd_con!estrnro & " AND " & _
                 " (htetdesde <= " & ConvFecha(FechaDeFin) & ") AND " & _
                 " ((" & ConvFecha(FechaDeFin) & " <= htethasta) or (htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        
        If rs_Estructura.EOF Then
            TipoDia_Ok = False
        End If
    End If

    If CBool(TipoDia_Ok) Then
        StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
                 " AND tdnro =" & rs_tipd_con!tdnro & _
                 " AND elfechadesde <=" & ConvFecha(FechaDeFin) & _
                 " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
        OpenRecordset StrSql, rs_Lic
        
        dias = 0
        Do While Not rs_Lic.EOF
            dias = CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
            SumaDias = SumaDias + dias
            
            'Marco la licencia para que no se pueda Borrar
            StrSql = "UPDATE emp_lic SET pronro = " & NroProc & _
                     " WHERE emp_licnro = " & rs_Lic!emp_licnro
            objConn.Execute StrSql, , adExecuteNoRecords
            
            rs_Lic.MoveNext
        Loop
    End If
    rs_tipd_con.MoveNext
Loop

If Month(FechaDeInicio) = 2 Then 'Febrero
    If Biciesto(Year(FechaDeInicio)) Then
        If SumaDias >= 29 Then
            Valor = 30
        Else
            Valor = SumaDias
        End If
    Else
        If SumaDias >= 28 Then
            Valor = 30
        Else
            Valor = SumaDias
        End If
    End If
Else
    If SumaDias > 30 Then
        Valor = 30
    Else
        Valor = SumaDias
    End If
End If

Valor = SumaDias
Bien = True
    

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

If rs_Lic.State = adStateOpen Then rs_Lic.Close
Set rs_Lic = Nothing

If rs_tipd_con.State = adStateOpen Then rs_tipd_con.Close
Set rs_tipd_con = Nothing

If rs_PeriodoGTI.State = adStateOpen Then rs_PeriodoGTI.Close
Set rs_PeriodoGTI = Nothing
End Sub


Public Sub bus_Vales()
' ---------------------------------------------------------------------------------------------
' Descripcion: Dtos de Vales
' Autor      : FGZ
' Fecha      : 14/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Opcion As Integer   '1 - Primera Quincena
                        '2 - Segunda Quincena
                        '3 - Mensual
Dim FechaDeInicio As Date
Dim FechaDeFin As Date

'Dim Param_cur As New ADODB.Recordset
Dim rs_Vales As New ADODB.Recordset
Dim PrimerDia As Integer
Dim UltimoDia As Integer

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Opcion = Arr_Programa(NroProg).Auxint1
    Else
        Exit Sub
    End If


Select Case Opcion
Case 1: 'Primera Quincena
    PrimerDia = 1
    UltimoDia = 15
    
    StrSql = "SELECT * FROM vales " & _
             " WHERE empleado = " & buliq_empleado!ternro & _
             " AND pliqdto = " & buliq_periodo!PliqNro & _
             " AND (pronro is null OR pronro = 0) " & _
             " AND valfecped >= " & ConvFecha(CDate(PrimerDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde))) & _
             " AND valfecped <= " & ConvFecha(CDate(UltimoDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde)))
    OpenRecordset StrSql, rs_Vales

    Do While Not rs_Vales.EOF
'        If 1 <= Day(rs_Vales!valfecped) And Day(rs_Vales!valfecped) <= 15 Then
            Valor = Valor + rs_Vales!valmonto
'        End If
        
        rs_Vales.MoveNext
    Loop

    'Actualiza Vales (Marca el vale)
    StrSql = "UPDATE vales SET pronro = " & NroProc & _
             " WHERE empleado = " & buliq_empleado!ternro & _
             " AND pliqdto = " & buliq_periodo!PliqNro & _
             " AND (pronro is null OR pronro = 0) " & _
             " AND valfecped >= " & ConvFecha(CDate(PrimerDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde))) & _
             " AND valfecped <= " & ConvFecha(CDate(UltimoDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde)))
    objConn.Execute StrSql, , adExecuteNoRecords
    
    
Case 2: 'Segunda Quincena
    PrimerDia = 16
    If Month(buliq_periodo!pliqdesde) = 12 Then
        UltimoDia = 31
    Else
        UltimoDia = Day(CDate("01/" & (Month(buliq_periodo!pliqdesde) + 1) & "/" & Year(buliq_periodo!pliqdesde)) - 1)
    End If
    
    StrSql = "SELECT * FROM vales " & _
             " WHERE empleado = " & buliq_empleado!ternro & _
             " AND pliqdto = " & buliq_periodo!PliqNro & _
             " AND (pronro is null OR pronro = 0) " & _
             " AND valfecped >= " & ConvFecha(CDate(PrimerDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde))) & _
             " AND valfecped <= " & ConvFecha(CDate(UltimoDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde)))
    OpenRecordset StrSql, rs_Vales

    Do While Not rs_Vales.EOF
'        If 16 <= Day(rs_Vales!valfecped) Then
            Valor = Valor + rs_Vales!valmonto
'        End If
        
        rs_Vales.MoveNext
    Loop
    
    'Actualiza los Vales (Marca el vale con el nro de proceso de Liquiacion)
    StrSql = "UPDATE vales SET pronro = " & NroProc & _
             " WHERE empleado = " & buliq_empleado!ternro & _
             " AND pliqdto = " & buliq_periodo!PliqNro & _
             " AND (pronro is null OR pronro = 0) " & _
             " AND valfecped >= " & ConvFecha(CDate(PrimerDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde))) & _
             " AND valfecped <= " & ConvFecha(CDate(UltimoDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde)))
    objConn.Execute StrSql, , adExecuteNoRecords
    
Case 3: 'Mensual
    PrimerDia = 1
    If Month(buliq_periodo!pliqdesde) = 12 Then
        UltimoDia = 31
    Else
        UltimoDia = Day(CDate("01/" & (Month(buliq_periodo!pliqdesde) + 1) & "/" & Year(buliq_periodo!pliqdesde)) - 1)
    End If

    StrSql = "SELECT * FROM vales " & _
             " WHERE empleado = " & buliq_empleado!ternro & _
             " AND pliqdto = " & buliq_periodo!PliqNro & _
             " AND (pronro is null OR pronro = 0) " & _
             " AND valfecped >= " & ConvFecha(CDate(PrimerDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde))) & _
             " AND valfecped <= " & ConvFecha(CDate(UltimoDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde)))
    OpenRecordset StrSql, rs_Vales

    Do While Not rs_Vales.EOF
        Valor = Valor + rs_Vales!valmonto
       
        rs_Vales.MoveNext
    Loop
    
    'Actualiza los Vales (Marca el vale con el nro de proceso de Liquiacion)
    StrSql = "UPDATE vales SET pronro = " & NroProc & _
             " WHERE empleado = " & buliq_empleado!ternro & _
             " AND pliqdto = " & buliq_periodo!PliqNro & _
             " AND (pronro is null OR pronro = 0) " & _
             " AND valfecped >= " & ConvFecha(CDate(PrimerDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde))) & _
             " AND valfecped <= " & ConvFecha(CDate(UltimoDia & "/" & Month(buliq_periodo!pliqdesde) & "/" & Year(buliq_periodo!pliqdesde)))
    objConn.Execute StrSql, , adExecuteNoRecords
    
End Select
Bien = True
    

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Vales.State = adStateOpen Then rs_Vales.Close
Set rs_Vales = Nothing

End Sub



Public Sub bus_Constantes()
' ---------------------------------------------------------------------------------------------
' Descripcion: Valores Constantes
' Autor      : FGZ
' Fecha      : 20/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

'Dim Param_cur As New ADODB.Recordset

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Valor = Arr_Programa(NroProg).Auxint1
    Else
        Exit Sub
    End If

    Bien = True

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

End Sub



Public Sub bus_ParametroConcepto()
' ---------------------------------------------------------------------------------------------
' Descripcion: Parametro de un concepto
' Autor      : FGZ
' Fecha      : 22/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim rs_wf_tpa As New ADODB.Recordset

    Bien = False
    Valor = 0
   
    StrSql = "SELECT * FROM " & TTempWF_tpa & _
             " ORDER BY ftorden DESC"
    OpenRecordset StrSql, rs_wf_tpa
    
    If Not rs_wf_tpa.EOF Then
        Valor = rs_wf_tpa!Valor
        Bien = True
    End If

' Cierro todo y libero
If rs_wf_tpa.State = adStateOpen Then rs_wf_tpa.Close
Set rs_wf_tpa = Nothing

End Sub


Public Sub bus_SueldoRemun()
' ---------------------------------------------------------------------------------------------
' Descripcion: Remuneracion del Empleado
' Autor      : FGZ
' Fecha      : 27/05/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Oblig As Boolean        ' Obligatorio retornar valor

'Dim Param_cur As New ADODB.Recordset
    
    Bien = False
    Valor = 0

    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Oblig = CBool(Arr_Programa(NroProg).Auxlog1)
    Else
        Exit Sub
    End If
    
    If Not EsNulo(buliq_empleado!empremu) Then
        Valor = buliq_empleado!empremu
        Bien = True
    Else
        If Not Oblig Then
            Bien = True
        End If
    End If
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing
    
End Sub


Public Sub bus_Sis_MesActual()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busqueda de sistema. Mes de la Liquidacion Actual
' Autor      : FGZ
' Fecha      : 03/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Bien = False
    Valor = 0
   
    If Not EsNulo(buliq_periodo!pliqmes) Then
        Valor = buliq_periodo!pliqmes
        Bien = True
    End If
End Sub


Public Sub bus_Sis_AnioActual()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busqueda de sistema. A�o de la Liquidacion Actual
' Autor      : FGZ
' Fecha      : 03/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Bien = False
    Valor = 0
   
    If Not EsNulo(buliq_periodo!pliqanio) Then
        Valor = buliq_periodo!pliqanio
        Bien = True
    End If
End Sub

Public Sub bus_Sis_SemestreActual()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busqueda de sistema. Semestre de la Liquidacion Actual
' Autor      : FGZ
' Fecha      : 03/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Bien = False
    Valor = 0
   
    If Not EsNulo(buliq_periodo!pliqmes) Then
        Select Case buliq_periodo!pliqmes
        Case 1, 2, 3, 4, 5, 6:
            Valor = 1
        Case 7, 8, 9, 10, 11, 12:
            Valor = 2
        Case Else
        End Select
        Bien = True
    End If
End Sub


Public Sub bus_Sis_ModeloLiqActual()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busqueda de sistema. Modelo de la Liquidacion Actual
' Autor      : FGZ
' Fecha      : 05/02/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Bien = False
    Valor = 0
   
    If Not EsNulo(buliq_proceso!tprocnro) Then
        Valor = buliq_proceso!tprocnro
        Bien = True
    End If
End Sub


Public Sub bus_Sis_Dias_Mes()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busqueda de sistema. Dias del Mes
' Autor      : FGZ
' Fecha      : 04/03/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Bien = False
    Valor = 0
   
    If buliq_periodo!pliqmes < 12 Then
        Valor = DateDiff("d", "01/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio, "01/" & buliq_periodo!pliqmes + 1 & "/" & buliq_periodo!pliqanio)
    Else
        Valor = DateDiff("d", "01/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio, "01/01/" & buliq_periodo!pliqanio + 1)
    End If
    Bien = True
End Sub


Public Sub bus_DiasVac()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer
Dim q As Integer

Dim ValorCoord As Single
Dim Encontro As Boolean
Dim ternro As Long

Dim DiasProporcion As Integer
Dim FactorDivision As Integer

Dim NroVac As Long
Dim cantdias As Integer
Dim Columna As Integer
Dim NroGrilla As Long
'Dim Param_cur As New ADODB.Recordset
Dim dias_trabajados As Integer

Dim FechaAux As Date
Dim Grilla_Ok As Boolean

    Bien = False
    Valor = 0
    ternro = buliq_empleado!ternro
    
'    'Busco los parametros
'    ' Obtener los parametros de la Busqueda
'    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        DiasProporcion = Arr_Programa(nroprog).auxint1
'    Else
'        Exit Sub
'    End If
    
    Call Politica(1501, Empleado_Fecha_Fin, Grilla_Ok)
    If Not Grilla_Ok Then
        Flog.writeline "Error cargando configuracion de la Politica 1501"
        Exit Sub
    Else
        DiasProporcion = st_CantidadDias
        FactorDivision = st_FactorDivision
        If FactorDivision = 0 Then
            FactorDivision = 1
        End If
    End If
    
    Call Politica(1502, Empleado_Fecha_Fin, Grilla_Ok)
    If Not Grilla_Ok Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    Else
        NroGrilla = st_Escala
    End If

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Exit Sub
    End If
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop


    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case ant
            Case 1:
                NroProg = rs_cabgrilla!grparnro_1
                Call bus_Anti0(antdia, antmes, antanio)
            Case 2:
                NroProg = rs_cabgrilla!grparnro_2
                Call bus_Anti0(antdia, antmes, antanio)
            Case 3:
                NroProg = rs_cabgrilla!grparnro_3
                Call bus_Anti0(antdia, antmes, antanio)
            Case 4:
                NroProg = rs_cabgrilla!grparnro_4
                Call bus_Anti0(antdia, antmes, antanio)
            Case 5:
                NroProg = rs_cabgrilla!grparnro_5
                Call bus_Anti0(antdia, antmes, antanio)
            End Select
            
            'OJO - Supuestamente este tipo de busqueda esta retornando el resultado en meses
            ' si esta busqueda no retorna meses, no va a encontrar el valor
            Parametros(j) = Valor 'Valor trae cantidad de meses
            
'            Call bus_Antiguedad("VACACIONES", CDate("31/12/" & buliq_periodo!pliqanio), antdia, antmes, antanio, q)
'            Parametros(j) = (antanio * 12) + antmes
        Case Else:
            Select Case j
            Case 1:
                NroProg = rs_cabgrilla!grparnro_1
                Call bus_Estructura
            Case 2:
                NroProg = rs_cabgrilla!grparnro_2
                Call bus_Estructura
            Case 3:
                NroProg = rs_cabgrilla!grparnro_3
                Call bus_Estructura
            Case 4:
                NroProg = rs_cabgrilla!grparnro_4
                Call bus_Estructura
            Case 5:
                NroProg = rs_cabgrilla!grparnro_5
                Call bus_Estructura
            End Select
            Parametros(j) = Valor
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
'        Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
'
'        If ValorCoord >= grilla_val(ant) Then
'            Call BusValor(7, Valor_Grilla, grilla_val, valor, Columna)
'            Bien = True
'        End If
'
'        rs_valgrilla.MoveNext
        Select Case ant
        Case 1:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
    
    If Not Encontro Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No encontr� valor en la escala "
        End If
    
        'Busco si existe algun valor para la estructura y ...
        'si hay que carga la columna correspondiente
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        StrSql = StrSql & " AND vgrvalor is not null"
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
        'StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        If Not rs_valgrilla.EOF Then
            Columna = rs_valgrilla!vgrorden
        Else
            Columna = 1
        End If
        
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
    
        If Parametros(ant) <= 6 Then

            'FactorDivision = 1
            If DiasProporcion = 20 Then
                If (dias_trabajados / DiasProporcion) / 7 * 5 > Fix((dias_trabajados / DiasProporcion) / 7 * 5) Then
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5) + 1
                Else
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5)
                End If
            Else
                cantdias = Round((dias_trabajados / DiasProporcion) / FactorDivision, 0)
            End If
            Valor = cantdias
            Bien = True
            
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Bien = False
        End If
    Else
        Valor = cantdias
        Bien = True
    End If
   
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub bus_DiasVac2()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de Diascorr.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_Vacacion As New ADODB.Recordset
Dim rs_vacdiascor As New ADODB.Recordset
    
    Valor = 0
    Bien = False
    
    StrSql = " SELECT * FROM vacacion WHERE vacacion.vacanio = " & buliq_periodo!pliqanio
    OpenRecordset StrSql, rs_Vacacion
    If Not rs_Vacacion.EOF Then
        StrSql = "SELECT * FROM vacdiascor WHERE vacnro = " & rs_Vacacion!vacnro & " AND Ternro = " & buliq_empleado!ternro
        OpenRecordset StrSql, rs_vacdiascor
        If Not rs_vacdiascor.EOF Then
            Valor = rs_vacdiascor!vdiascorcant
            Bien = True
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Na hoy dias correspondientes generados para el periodo"
            End If
            Exit Sub
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No existe el periodo de vacaciones"
        End If
        Exit Sub
    End If
    
   
    
' Cierro todo y libero
If rs_Vacacion.State = adStateOpen Then rs_Vacacion.Close
If rs_vacdiascor.State = adStateOpen Then rs_vacdiascor.Close

Set rs_Vacacion = Nothing
Set rs_vacdiascor = Nothing
End Sub

Public Sub bus_DiasVac_Masivos()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer
Dim q As Integer

Dim ValorCoord As Single
Dim Encontro As Boolean
Dim ternro As Long

Dim DiasProporcion As Integer
Dim FactorDivision As Integer

Dim NroVac As Long
Dim cantdias As Integer
Dim Columna As Integer
Dim NroGrilla As Long
'Dim Param_cur As New ADODB.Recordset
Dim dias_trabajados As Integer

    Bien = False
    Valor = 0
    ternro = buliq_empleado!ternro
    
'    'Busco los parametros
'    ' Obtener los parametros de la Busqueda
'    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        DiasProporcion = Arr_Programa(nroprog).auxint1
'    Else
'        Exit Sub
'    End If
    
    ' Maxi: Mal se debe sacar de la politica de GTI
    StrSql = "SELECT * FROM tipdia WHERE tdnro = 2 " '2 es vacaciones
    OpenRecordset StrSql, objRs
    If Not objRs.EOF Then
        NroGrilla = objRs!tdgrilla
        tdinteger3 = objRs!tdinteger3

        If tdinteger3 <> 20 And tdinteger3 <> 365 And tdinteger3 <> 360 Then
            'El campo auxiliar3 del Tipo de D�a para Vacaciones no est� configurado para Proporcionar la cant. de d�as de Vacaciones.
            Exit Sub
        End If
        DiasProporcion = tdinteger3
    End If

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Exit Sub
    End If
    
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop


    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Call bus_Antiguedad("VACACIONES", CDate("31/12/" & buliq_periodo!pliqanio), antdia, antmes, antanio, q)
            Parametros(j) = (antanio * 12) + antmes
        Case Else:
            Select Case j
            Case 1:
                NroProg = rs_cabgrilla!grparnro_1
                Call bus_Estructura
            Case 2:
                NroProg = rs_cabgrilla!grparnro_2
                Call bus_Estructura
            Case 3:
                NroProg = rs_cabgrilla!grparnro_3
                Call bus_Estructura
            Case 4:
                NroProg = rs_cabgrilla!grparnro_4
                Call bus_Estructura
            Case 5:
                NroProg = rs_cabgrilla!grparnro_5
                Call bus_Estructura
            End Select
            Parametros(j) = Valor
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
'        Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
'
'        If ValorCoord >= grilla_val(ant) Then
'            Call BusValor(7, Valor_Grilla, grilla_val, valor, Columna)
'            Bien = True
'        End If
'
'        rs_valgrilla.MoveNext
        Select Case ant
        Case 1:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
    
    If Not Encontro Then
    
        'Busco si existe algun valor para la estructura y ...
        'si hay que carga la columna correspondiente
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        StrSql = StrSql & " AND vgrvalor is not null"
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
        'StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        If Not rs_valgrilla.EOF Then
            Columna = rs_valgrilla!vgrorden
        Else
            Columna = 1
        End If
        
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
    
        If Parametros(ant) <= 6 Then

            FactorDivision = 1
            If DiasProporcion = 20 Then
                If (dias_trabajados / DiasProporcion) / 7 * 5 > Fix((dias_trabajados / DiasProporcion) / 7 * 5) Then
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5) + 1
                Else
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5)
                End If
            Else
                cantdias = Round((dias_trabajados / DiasProporcion) / FactorDivision, 0)
            End If
            Valor = cantdias
            Bien = True
            
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Bien = False
        End If
    Else
        Valor = cantdias
        Bien = True
    End If
   
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub



Public Sub bus_DiasMesCalendario_enEstructura()
' ---------------------------------------------------------------------------------------------
' Descripcion: Dias en el tipo de Estructura en el mes que se esta liquidando
' Autor      : FGZ
' Fecha      : 06/04/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TipoEstr As Long      ' Tipo de Estructura
                            
'Dim Param_cur As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Desde As Date
Dim Hasta As Date

Dim dia As Integer
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoEstr = Arr_Programa(NroProg).Auxint1
    Else
        Exit Sub
    End If

FechaDesde = buliq_periodo!pliqdesde
FechaHasta = buliq_periodo!pliqhasta

If Not CBool(buliq_empleado!empest) Then
    If FechaHasta > Empleado_Fecha_Fin Then
        FechaHasta = Empleado_Fecha_Fin
    End If
End If

'Busco de estructura
StrSql = " SELECT htetdesde,htethasta FROM his_estructura " & _
         " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
         " tenro =" & TipoEstr & " AND " & _
         " (htetdesde <= " & ConvFecha(FechaHasta) & ") AND " & _
         " ((" & ConvFecha(FechaDesde) & " <= htethasta) or (htethasta is null))" & _
         " ORDER BY htetdesde"
OpenRecordset StrSql, rs_Estructura

dia = 0
Do While Not rs_Estructura.EOF
    If FechaDesde < rs_Estructura!htetdesde Then
        Desde = rs_Estructura!htetdesde
    Else
        Desde = FechaDesde
    End If
    
    If EsNulo(rs_Estructura!htethasta) Then
        Hasta = FechaHasta
    Else
        If rs_Estructura!htethasta < FechaHasta Then
            Hasta = rs_Estructura!htethasta
        Else
            Hasta = FechaHasta
        End If
    End If
    
    'dia = DateDiff("d", Desde, Hasta)
    Call DIF_FECHAS2(Desde, Hasta, aux1, aux2, aux3)
    dia = dia + (aux1 + 1)
    
    rs_Estructura.MoveNext
Loop

Valor = dia
Bien = True
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

End Sub




Public Sub bus_AsignacionesFliares()
' ---------------------------------------------------------------------------------------------
' Descripcion: Asignaciones Familiares
' Autor      : FGZ
' Fecha      : 15/04/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Incapacitado As Boolean         '-1 (True) / 0 (False)
Dim Edad As Integer                 'cant de a�os o nulo o vacio
Dim sexo As Integer                 '1(Masc) / 2 (Fem) / 3 (Todos)
Dim Estudia As Integer              '1(si) / 2 (no) / 3 (indefinido)
Dim Ayuda_Escolar As Integer        '1(si) / 2 (no) / 3 (indefinido)
Dim Suma_FliaNumerosa As Integer    'Siempre viene 1. No se usa mas
Dim Paga_FliaNumerosa As Integer    'siempre viene 1. No se usa mas
Dim Trabaja_Conyuge As Integer      '1(si) / 2 (no) / 3 (no importa)
Dim Retroactivo_Prenatal As Boolean '-1(TRUE) / 0(FALSE)
Dim Nivel_Estudio As String         'nivnro,nivnro,....
Dim Periodo_Escolar As Integer      'nro del periodo escolar
Dim Parentesco As Integer           'codigo del parentesco
                            
Dim Fam_niv_est     As Integer
Dim Fam_peri_escol  As Integer
Dim Fam_estudia     As Integer
Dim Fecha_vto_asig  As Date
Dim Fin_periodo_liq As Date
Dim Par_asig        As Integer

Dim suma_fn As Integer
Dim paga_fn As Integer
Dim edad_f As Integer
Dim conyuge As Integer
Dim pagaxhijo As Boolean
Dim niv_est_interno As Boolean
Dim interesa_estu As Boolean
Dim sexo_conyuge As Integer
Dim conyuge_trabaja As Integer

'Dim Param_cur As New ADODB.Recordset
Dim rs_PeriodoEsc As New ADODB.Recordset
Dim rs_Familiar As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim rs_estudio_actual As New ADODB.Recordset
Dim rs_Nivest As New ADODB.Recordset
  
'inicializo
Par_asig = 31
conyuge = False
pagaxhijo = False
suma_fn = 0
paga_fn = 0

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Incapacitado = Arr_Programa(NroProg).Auxint4
        Edad = Arr_Programa(NroProg).Auxint1
        sexo = Arr_Programa(NroProg).Auxint5
        Estudia = Arr_Programa(NroProg).Auxchar2
        Ayuda_Escolar = Arr_Programa(NroProg).Auxchar3
        Suma_FliaNumerosa = True
        Paga_FliaNumerosa = True
        Trabaja_Conyuge = Arr_Programa(NroProg).Auxchar5
        Retroactivo_Prenatal = Arr_Programa(NroProg).Auxint2
        Nivel_Estudio = IIf(EsNulo(Arr_Programa(NroProg).Auxchar1), 0, Arr_Programa(NroProg).Auxchar1)
        Periodo_Escolar = Arr_Programa(NroProg).Auxint3
        Parentesco = Arr_Programa(NroProg).Auxchar4
    Else
        Exit Sub
    End If

    ' VALIDAR SI AL EMPLEADO CORRESPONDE PAGARLE ASIGNACIONES FAMILIARES
    ' en funcion : dias trabajados en el mes
    
    'AYUDA ESCOLAR
    If Ayuda_Escolar = 1 Then
        StrSql = "SELECT * FROM edu_periodoesc WHERE edu_periodoesc.perescnro =" & Periodo_Escolar
        OpenRecordset StrSql, rs_PeriodoEsc
        If rs_PeriodoEsc.EOF Then
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Periodo Escolar Incorrecto."
            End If
            Exit Sub
        End If
    End If
        
    'FECHAS LIMITES DE VENCIMENTOS DE CERTIFICADOS
    Fecha_vto_asig = buliq_periodo!pliqdesde
    Fin_periodo_liq = buliq_periodo!pliqhasta


    StrSql = "SELECT * FROM familiar INNER JOIN tercero ON tercero.ternro = familiar.ternro"
    StrSql = StrSql & " WHERE (familiar.empleado =" & buliq_empleado!ternro
    StrSql = StrSql & " AND familiar.parenro = " & Parentesco
    StrSql = StrSql & " AND familiar.famest = -1"
    StrSql = StrSql & " AND familiar.famsalario = -1)"
    StrSql = StrSql & " AND (familiar.famfecvto >=" & ConvFecha(Fecha_vto_asig) & " OR familiar.famfecvto is null)"
    StrSql = StrSql & " Order by tercero.terfecnac DESC"
    OpenRecordset StrSql, rs_Familiar
             
    Do While Not rs_Familiar.EOF
        conyuge = 3
        sexo_conyuge = 3
        conyuge_trabaja = 3
        
        If rs_Familiar!parenro = 2 Then 'hijo
            'calculo la edad del familiar
             edad_f = Calcular_Edad(rs_Familiar!terfecnac)
             
            conyuge = 3
            sexo_conyuge = 3
            conyuge_trabaja = IIf(CBool(rs_Familiar!famtrab), 1, 2)
            'conyuge_trabaja = Trabaja_Conyuge
        End If
                             
        'en los conyuges no interesa si estudia o no
        interesa_estu = True  'el default es que interese
        
        If rs_Familiar!parenro = 3 Then 'conyuge
            sexo_conyuge = IIf(CBool(rs_Familiar!tersex), 1, 2)
            conyuge_trabaja = IIf(CBool(rs_Familiar!famtrab), 1, 2)
            If (rs_Familiar!tersex) And (rs_Familiar!faminc) And (Not rs_Familiar!famtrab) Then
                conyuge = False
            Else
                If (Not rs_Familiar!tersex) And rs_Familiar!famtrab Then
                    conyuge = True
                Else
                    conyuge = False
                End If
            End If
        
            If (sexo_conyuge = 1 And conyuge_trabaja = 1) Then 'Conyuge masculino y trabaja
               GoTo SiguienteFamiliar
            End If
            interesa_estu = False
            edad_f = 0
        End If
        
        'buscar el nivel de estudio
        Fam_estudia = False
        StrSql = "SELECT * FROM estudio_actual WHERE ternro = " & rs_Familiar!ternro
        OpenRecordset StrSql, rs_estudio_actual
        If Not rs_estudio_actual.EOF Then
            If Not EsNulo(rs_estudio_actual!nivnro) Then
                StrSql = "SELECT * FROM nivest WHERE nivnro =" & rs_estudio_actual!nivnro
                OpenRecordset StrSql, rs_Nivest
                If Not rs_Nivest.EOF Then
                    niv_est_interno = rs_Nivest!nivsist
                    Fam_niv_est = rs_Nivest!nivnro
                    Fam_estudia = IIf(EsNulo(rs_estudio_actual!nivnro), 2, 1)
                Else
                    niv_est_interno = False
                End If
            End If
        End If
            
            
        'ACA SE PODRIA VALIDAR LA VALIDEZ DEL CERTIFICADO ESCOLAR
        'INICIO
        'FIN
        'SI NO ES VALIDO EL CERTIFICADO, ASIGNAR A FAM-NIV-EST = ?
        'FAM-ESTUDIA = (estudio_actual.nivnro <> ?)
                   
        If (CBool(Incapacitado) = CBool(rs_Familiar!faminc)) Then
            If (Parentesco = rs_Familiar!parenro) And (edad_f <= Edad Or EsNulo(Edad)) And _
                ((sexo = 1 And CBool(rs_Familiar!tersex)) Or (sexo = 2 And Not CBool(rs_Familiar!tersex)) Or sexo = 3) Then
                If (conyuge_trabaja = Trabaja_Conyuge Or Trabaja_Conyuge = 3) Then
                    If (Estudia = Fam_estudia Or Estudia = 3 Or (Not interesa_estu)) Then
                        If (InStr(1, Nivel_Estudio, Fam_niv_est) <> 0 Or EsNulo(Nivel_Estudio) Or (Not niv_est_interno)) Then
                            If (((Ayuda_Escolar = 1) And InStr(1, Nivel_Estudio, Fam_niv_est) <> 0) Or ((Ayuda_Escolar = 3) And _
                                    Not Retroactivo_Prenatal)) Then
                                Valor = Valor + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
       
SiguienteFamiliar:
        rs_Familiar.MoveNext
    Loop
    
    Bien = True
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_estudio_actual.State = adStateOpen Then rs_estudio_actual.Close
Set rs_estudio_actual = Nothing

If rs_PeriodoEsc.State = adStateOpen Then rs_PeriodoEsc.Close
Set rs_PeriodoEsc = Nothing

If rs_Familiar.State = adStateOpen Then rs_Familiar.Close
Set rs_Familiar = Nothing

If rs_Nivest.State = adStateOpen Then rs_Nivest.Close
Set rs_Nivest = Nothing

If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
Set rs_Tercero = Nothing

End Sub



Public Function Calcular_Edad(ByVal Fecha As Date) As Integer
'...........................................................................
' Archivo       : edad.i                              fecha ini. : 20/01/92
' Nombre progr. :
' tipo programa : FGZ
' Descripcion   :
'...........................................................................
Dim a�os  As Integer

    a�os = Year(Date) - Year(Fecha)
    If Month(Date) < Month(Fecha) Then
       a�os = a�os - 1
    Else
        If Month(Date) = Month(Fecha) Then
            If Day(Date) < Day(Fecha) Then
                a�os = a�os - 1
            End If
        End If
    End If
    Calcular_Edad = a�os
End Function


Public Sub bus_DiasHabilesMesLiquidacion()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias Habiles entre dos fechas
' Autor      : FGZ
' Fecha      : 24/05/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim AnteriorPosterior As Integer    '1- meses para atras
                                    '2- meses para adelante
Dim CantMeses As Integer            'Cantidad de meses
Dim DiasHabiles As Single
Dim dia As Date
Dim EsFeriado As Boolean
Dim IncluyeSabados As Boolean       'si se cuentan los sabados
Dim porcentaje As Integer           'porcentaje de los sabados incluidos

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim objFeriado As New Feriado

'Dim Param_cur As New ADODB.Recordset

    
' inicializacion de variables
    Set objFeriado.Conexion = objConn
    'Set objFeriado.ConexionTraza = objConn

    Bien = False
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        AnteriorPosterior = Arr_Programa(NroProg).Auxint1
        CantMeses = Arr_Programa(NroProg).Auxint2
        porcentaje = Arr_Programa(NroProg).Auxint3
        IncluyeSabados = CBool(Arr_Programa(NroProg).Auxlog1)
    Else
        Exit Sub
    End If

    Select Case AnteriorPosterior
    Case 1:     'para atras
        FechaDesde = DateAdd("m", -1 * CantMeses, buliq_periodo!pliqdesde)
        FechaHasta = DateAdd("m", 1, FechaDesde) - 1
    Case 2:     'para delante
        FechaDesde = DateAdd("m", CantMeses, buliq_periodo!pliqdesde)
        FechaHasta = DateAdd("m", 1, FechaDesde) - 1
    Case Else   'para delante
        FechaDesde = DateAdd("m", CantMeses, buliq_periodo!pliqdesde)
        FechaHasta = DateAdd("m", 1, FechaDesde) - 1
    End Select
    
dia = FechaDesde
Do While dia <= FechaHasta
    
    EsFeriado = objFeriado.Feriado(dia, buliq_empleado!ternro, HACE_TRAZA)
    
    If Not EsFeriado And Not Weekday(dia) = 1 Then
        ' No es feriado no Domingo
        If Weekday(dia) = 7 Then 'Sabado
            If IncluyeSabados Then
                DiasHabiles = DiasHabiles + (1 * porcentaje / 100)
            End If
        Else
            DiasHabiles = DiasHabiles + 1
        End If
    End If
    dia = dia + 1
Loop

Bien = True
Valor = DiasHabiles

End Sub

Public Sub bus_PromedioVacaciones()
' ---------------------------------------------------------------------------------------------
' Descripcion: Promedio de Vacaciones
' Autor      : FGZ
' Fecha      : 27/05/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Anual As Boolean        'TRUE - Anual
                            'FALSE - Semestral
Dim MesActual As Boolean    'TRUE - hasta mes actual
                            'FALSE - hasta mes anterior
Dim Opcion As Integer       ' 0 - Pago de adelanto
                            ' 1 - Descuento de adelanto
                            ' 2 - Descuento de vacaciones
                            ' 3 - Descuento de vacaciones
Dim PromedioSin0 As Boolean 'TRUE - Promedio sin 0
                            'FALSE - Promedio con 0
Dim Acumulador As Long

'Dim Param_cur As New ADODB.Recordset
Dim rs_Vacpagdesc As New ADODB.Recordset
Dim rs_Vacacion As New ADODB.Recordset
Dim rs_Acu_Mes As New ADODB.Recordset

Dim MesInicio As Integer
Dim MesFin As Integer
Dim Anio As Integer
Dim Suma As Single
Dim Cantidad As Single
Dim esMonto As Boolean
   
    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Opcion = Arr_Programa(NroProg).Auxint4
        Anual = IIf(Arr_Programa(NroProg).Auxint2 = 0, True, False)
        MesActual = IIf(Arr_Programa(NroProg).Auxint3 = 0, True, False)
        PromedioSin0 = IIf(Arr_Programa(NroProg).Auxint5 = 0, True, False)
        Acumulador = Arr_Programa(NroProg).Auxint1
        esMonto = CBool(Arr_Programa(NroProg).Auxlog1)
    Else
        Exit Sub
    End If

    'seteo desde y hasta
    Anio = buliq_periodo!pliqanio
    If Anual Then
        MesInicio = 1
    Else
        MesInicio = 6
    End If
    If MesActual Then
        MesFin = buliq_periodo!pliqmes
    Else
        MesFin = buliq_periodo!pliqmes - 1
    End If
    
    StrSql = " SELECT * FROM vacpagdesc " & _
             " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
             " pliqnro =" & buliq_proceso!PliqNro & " AND " & _
             " tprocnro =" & buliq_proceso!tprocnro & " AND " & _
             " pago_dto =" & Opcion
    OpenRecordset StrSql, rs_Vacpagdesc

    If Not rs_Vacpagdesc.EOF Then
        StrSql = " SELECT * FROM vacacion WHERE vacacion.vacnro = " & rs_Vacpagdesc!vacnro
        OpenRecordset StrSql, rs_Vacacion
        If Not rs_Vacacion.EOF Then
            If rs_Vacacion!vacanio <> buliq_periodo!pliqanio Then
                Anio = buliq_periodo!pliqanio - 1
                If Anual Then
                    MesInicio = 1
                    MesFin = 12
                Else
                    MesInicio = 7
                    MesFin = 12
                End If
            End If
            
            StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                     " AND acunro =" & Acumulador & _
                     " AND " & Anio & " = amanio " & _
                     " AND (ammes >=" & MesInicio & " AND " & MesFin & " >= ammes)"
            OpenRecordset StrSql, rs_Acu_Mes
            
            Suma = 0
            Cantidad = 0
            Do While Not rs_Acu_Mes.EOF
                If Not PromedioSin0 Then
                    If esMonto Then
                        Suma = Suma + rs_Acu_Mes!ammonto
                    Else
                        Suma = Suma + rs_Acu_Mes!amcant
                    End If
                    Cantidad = Cantidad + 1
                Else
                    If esMonto Then
                        If rs_Acu_Mes!ammonto <> 0 Then
                            Suma = Suma + rs_Acu_Mes!ammonto
                            Cantidad = Cantidad + 1
                        End If
                    Else
                        If rs_Acu_Mes!amcant <> 0 Then
                            Suma = Suma + rs_Acu_Mes!amcant
                            Cantidad = Cantidad + 1
                        End If
                    End If
                End If
                rs_Acu_Mes.MoveNext
            Loop
            
            If Cantidad <> 0 Then
                Valor = Suma / Cantidad
            End If
            
        End If
        
        Bien = True
    End If
        
    
' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Vacpagdesc.State = adStateOpen Then rs_Vacpagdesc.Close
Set rs_Vacpagdesc = Nothing

If rs_Vacacion.State = adStateOpen Then rs_Vacacion.Close
Set rs_Vacacion = Nothing

If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
Set rs_Acu_Mes = Nothing


End Sub

Public Sub bus_Antfases(ByVal A_fecha As Date, ByVal D_Fecha As Date, ByRef dia As Integer, ByRef Mes As Integer, ByRef Anio As Integer)
' ---------------------------------------------------------------------------------------------
' Descripcion: Calcula la antiguedad a una fecha de un empleado en :
'              en dias, meses y a�os
' Autor      : FGZ
' Fecha      : 08/06/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim aux1 As Integer
Dim aux2 As Integer
Dim aux3 As Integer
Dim fecalta As Date
Dim fecbaja As Date
Dim seguir As Date
Dim q As Integer

Dim rs_Fases As New ADODB.Recordset

StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro
OpenRecordset StrSql, rs_Fases

'StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
'         " AND " & NombreCampo & " = -1 " & _
'         " AND not altfec is null " & _
'         " AND not (bajfec is null AND estado = 0)" & _
'         " AND altfec <= " & ConvFecha(fecha_fin)
'OpenRecordset StrSql, rs_Fases

Do While Not rs_Fases.EOF
    fecalta = rs_Fases!altfec
    
    'Verificar si se trata de un registro completo (alta/baja) o solo de un alta
    If EsNulo(rs_Fases!bajfec) Then
        fecbaja = A_fecha 'solo es un alta, tomar el a-fecha
    Else
        fecbaja = rs_Fases!bajfec 'se trata de un registro completo
    End If
                           
    If Not EsNulo(fecbaja) Then  'ya esta dado de baja
        If fecalta < D_Fecha And fecbaja < D_Fecha Then GoTo siguiente
        If fecalta < D_Fecha And fecbaja > D_Fecha Then fecalta = D_Fecha
        
        Call DIF_FECHAS2(fecalta, fecbaja, aux1, aux2, aux3)
        dia = dia + aux1
        Mes = Mes + aux2 + Int(dia / 30)
        Anio = Anio + aux3 + Int(Mes / 12)
        dia = dia Mod 30
        Mes = Mes Mod 12
     
     End If
siguiente:
    rs_Fases.MoveNext
Loop

' Cierro todo y Libero
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing

End Sub



Public Sub bus_DiasSAC()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias del semestre actual para SAC
' Autor      : FGZ
' Fecha      : 10/06/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Fec_Ini_Sem     As Date
Dim Fec_Fin_Sem     As Date
Dim Fec_Ini_1_Sem   As Date
Dim Fec_Ini_2_Sem   As Date
Dim Fec_Ini_Calc    As Date
Dim Fec_Fin_Calc    As Date
Dim Fec_Fin_1_Sem   As Date
Dim Fec_Fin_2_Sem   As Date
Dim Dias_Sac        As Single
Dim Dias_Aus        As Single

Dim A_fecha        As Date
Dim Maximo         As Integer
Dim Tolerancia     As Integer
Dim TiposDeDia     As String

'Dim Param_cur As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

    Bien = False
    Valor = 0
    'A_fecha = buliq_periodo!pliqdesde
    A_fecha = buliq_periodo!pliqhasta
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Maximo = Arr_Programa(NroProg).Auxint1
        Tolerancia = Arr_Programa(NroProg).Auxint2
        TiposDeDia = IIf(Not EsNulo(Arr_Programa(NroProg).Auxchar), Arr_Programa(NroProg).Auxchar, " ")
    Else
        Exit Sub
    End If


    'calculo de inicio del semetre
    Fec_Ini_1_Sem = CDate("01/01/" & Year(A_fecha))
    Fec_Ini_2_Sem = CDate("01/07/" & Year(A_fecha))
    Fec_Ini_Sem = IIf(A_fecha >= Fec_Ini_2_Sem, Fec_Ini_2_Sem, Fec_Ini_1_Sem)
    Fec_Fin_1_Sem = CDate("30/06/" & Year(A_fecha))
    Fec_Fin_2_Sem = CDate("31/12/" & Year(A_fecha))
    Fec_Fin_Sem = IIf(A_fecha >= Fec_Ini_2_Sem, Fec_Fin_2_Sem, Fec_Fin_1_Sem)
    Fec_Fin_Sem = IIf(A_fecha < Fec_Fin_Sem, A_fecha, Fec_Fin_Sem)
    ' SE AGREGARON ESTAS 2 INICIALIZACIONES
    Fec_Ini_Calc = Fec_Ini_Sem
    Fec_Fin_Calc = Fec_Fin_Sem


    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE estado = -1 AND real = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If rs_Fases!altfec > Fec_Ini_Sem Then
            Fec_Ini_Calc = rs_Fases!altfec
        Else
            Fec_Ini_Calc = Fec_Ini_Sem
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna fase activa para el empleado " & buliq_empleado!empleg
        End If
                
        Bien = False
        Exit Sub
    End If

    'Busco la ultima fase inactiva
    StrSql = "SELECT * FROM fases WHERE estado = 0 AND real = -1 AND empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If rs_Fases!bajfec < Fec_Ini_Sem Then
            Fec_Fin_Calc = Fec_Fin_Sem
        Else
            If rs_Fases!bajfec < Fec_Fin_Sem Then
                Fec_Fin_Calc = rs_Fases!bajfec
            Else
                Fec_Fin_Calc = Fec_Fin_Sem
            End If
        End If
    End If


    'Dias_Sac = Fec_Fin_Sem - Fec_Ini_Calc
    Dias_Sac = DateDiff("d", Fec_Ini_Calc, Fec_Fin_Calc)
    Dias_Sac = IIf(Dias_Sac >= (Maximo - Tolerancia), Maximo, Dias_Sac)

    'descontar las licencias en el semestre
'    StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
'             " AND (tdnro = 15 OR tdnro = 16 OR tdnro = 17 OR tdnro = 18 OR tdnro = 24)" & _
'             " AND elfechadesde <=" & ConvFecha(Fec_Fin_Sem) & _
'             " AND elfechahasta >= " & ConvFecha(Fec_Ini_Calc)

    If Not EsNulo(TiposDeDia) And TiposDeDia <> " " Then
        StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
                 " AND (tdnro IN (" & TiposDeDia & "))" & _
                 " AND elfechadesde <=" & ConvFecha(Fec_Fin_Sem) & _
                 " AND elfechahasta >= " & ConvFecha(Fec_Ini_Calc)
        OpenRecordset StrSql, rs_Lic
        
        Do While Not rs_Lic.EOF
            Dias_Aus = Dias_Aus + CantidadDeDias(Fec_Ini_Calc, Fec_Fin_Sem, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
            
            rs_Lic.MoveNext
        Loop
    Else
        Dias_Aus = 0
    End If

    Dias_Sac = Dias_Sac - Dias_Aus
    
    Valor = Dias_Sac
    Bien = True

End Sub

Public Sub bus_DiasSAC_Proporcional()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Dias del semestre actual para SAC con proporcion.
' Autor      : FGZ
' Fecha      : 10/06/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim Fec_Ini_Sem     As Date
Dim Fec_Fin_Sem     As Date
Dim Fec_Ini_1_Sem   As Date
Dim Fec_Ini_2_Sem   As Date
Dim Fec_Ini_Calc    As Date
Dim Fec_Fin_Calc    As Date
Dim Fec_Fin_1_Sem   As Date
Dim Fec_Fin_2_Sem   As Date
Dim Dias_Sac        As Single
Dim Dias_Aus        As Single

Dim A_fecha        As Date
Dim Maximo         As Integer
Dim Tolerancia     As Integer
Dim TiposDeDia     As String

'Dim Param_cur As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

    Bien = False
    Valor = 0
    'A_fecha = buliq_periodo!pliqdesde
    A_fecha = buliq_periodo!pliqhasta
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur

    If Arr_Programa(NroProg).Prognro <> 0 Then
        Maximo = Arr_Programa(NroProg).Auxint1
        Tolerancia = Arr_Programa(NroProg).Auxint2
        TiposDeDia = IIf(Not EsNulo(Arr_Programa(NroProg).Auxchar), Arr_Programa(NroProg).Auxchar, " ")
    Else
        Exit Sub
    End If


    'calculo de inicio del semetre
    Fec_Ini_1_Sem = CDate("01/01/" & Year(A_fecha))
    Fec_Ini_2_Sem = CDate("01/07/" & Year(A_fecha))
    Fec_Ini_Sem = IIf(A_fecha >= Fec_Ini_2_Sem, Fec_Ini_2_Sem, Fec_Ini_1_Sem)
    Fec_Fin_1_Sem = CDate("30/06/" & Year(A_fecha))
    Fec_Fin_2_Sem = CDate("31/12/" & Year(A_fecha))
    Fec_Fin_Sem = IIf(A_fecha >= Fec_Ini_2_Sem, Fec_Fin_2_Sem, Fec_Fin_1_Sem)
    Fec_Fin_Sem = IIf(A_fecha < Fec_Fin_Sem, A_fecha, Fec_Fin_Sem)
    ' SE AGREGARON ESTAS 2 INICIALIZACIONES
    Fec_Ini_Calc = Fec_Ini_Sem
    Fec_Fin_Calc = Fec_Fin_Sem



'    'Busco la ultima fase activa
'    StrSql = "SELECT * FROM fases WHERE estado = -1 AND real = -1 AND empleado = " & buliq_empleado!ternro
'    OpenRecordset StrSql, rs_Fases
'
'    If Not rs_Fases.EOF Then
'        rs_Fases.MoveLast
'        If rs_Fases!altfec > Fec_Ini_Sem Then
'            Fec_Ini_Calc = rs_Fases!altfec
'        Else
'            Fec_Ini_Calc = Fec_Ini_Sem
'        End If
'    Else
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna fase activa para el empleado " & buliq_empleado!empleg
'        End If
'
'        Bien = False
'        Exit Sub
'    End If

    'Busco la ultima fase inactiva
    StrSql = "SELECT * FROM fases WHERE real = -1 AND empleado = " & buliq_empleado!ternro
    StrSql = StrSql & " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        'Busco la fecha de Inicio
        If rs_Fases!altfec > Fec_Ini_Sem Then
            Fec_Ini_Calc = rs_Fases!altfec
        Else
            Fec_Ini_Calc = Fec_Ini_Sem
        End If
        'Busco la fecha de fin
        If Not EsNulo(rs_Fases!bajfec) Then
            Fec_Fin_Calc = rs_Fases!bajfec
        Else
            If Not EsNulo(buliq_empleado!empfbajaprev) Then
                Fec_Fin_Calc = buliq_empleado!empfbajaprev
            Else
                Fec_Fin_Calc = Fec_Fin_Sem
            End If
        End If
    Else
        Fec_Fin_Calc = Fec_Fin_Sem
    End If


    'Dias_Sac = Fec_Fin_Sem - Fec_Ini_Calc
    Dias_Sac = DateDiff("d", Fec_Ini_Calc, Fec_Fin_Calc) + 1
    Dias_Sac = IIf(Dias_Sac >= (Maximo - Tolerancia), Maximo, Dias_Sac)

    'descontar las licencias en el semestre
'    StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & buliq_empleado!ternro & " )" & _
'             " AND (tdnro = 15 OR tdnro = 16 OR tdnro = 17 OR tdnro = 18 OR tdnro = 24)" & _
'             " AND elfechadesde <=" & ConvFecha(Fec_Fin_Sem) & _
'             " AND elfechahasta >= " & ConvFecha(Fec_Ini_Calc)

    If Not EsNulo(TiposDeDia) And TiposDeDia <> " " Then
        StrSql = "SELECT * FROM emp_lic WHERE empleado = " & buliq_empleado!ternro & _
                 " AND tdnro IN (" & TiposDeDia & ") " & _
                 " AND elfechadesde <=" & ConvFecha(Fec_Fin_Sem) & _
                 " AND elfechahasta >= " & ConvFecha(Fec_Ini_Calc)
        OpenRecordset StrSql, rs_Lic
        
        Do While Not rs_Lic.EOF
            Dias_Aus = Dias_Aus + CantidadDeDias(Fec_Ini_Calc, Fec_Fin_Sem, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
            
            rs_Lic.MoveNext
        Loop
    Else
        Dias_Aus = 0
    End If

    Dias_Sac = Dias_Sac - Dias_Aus
    
    Valor = Dias_Sac
    Bien = True

End Sub


Public Sub bus_Vac_No_Gozadas(ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Vacaciones no Gozadas
' Autor      : FGZ
' Fecha      : 02/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fec_Fin As Date

Dim Maximo       As Single
Dim Tolerancia   As Single
Dim Inas_Ingreso As Single
Dim Diasvac     As Single
Dim Diasvactomados  As Single
Dim Genera       As Boolean
Dim Propor       As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin se toma a la fecha de baja de la ultima fase "
        Flog.writeline Espacios(Tabulador * 4) & "Si no esta dado de baja tomo la fecha de baja prevista y si "
        Flog.writeline Espacios(Tabulador * 4) & " si la fecha de baja prevista es nula tomo la fecha fin del proceso "
    End If
    
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND real = -1 " & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If Not EsNulo(rs_Fases!bajfec) Then
            Fec_Fin = rs_Fases!bajfec
        Else
            If Not EsNulo(buliq_empleado!empfbajaprev) Then
                Fec_Fin = buliq_empleado!empfbajaprev
            Else
                Fec_Fin = buliq_proceso!profecfin
            End If
        End If
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin es:  " & Fec_Fin
    End If
    
    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Maximo = Arr_Programa(NroProg).Auxint1
        Tolerancia = Arr_Programa(NroProg).Auxint2
        Inas_Ingreso = 0
    Else
        Exit Sub
    End If


    'A pedido de Analia
    ' run pronov/des01.p(empleado.ternro, FEC-fin, output diasvac,output genera).
    'Call bus_DiasVac_Masivos
    Call bus_DiasVac
    Diasvac = Valor
    
    If Not Bien Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se pudieron generar los dias de Vacaciones no Gozadas de: " & buliq_empleado!empleg
            Exit Sub
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Dias de vacaciones : " & Valor
        End If
    End If

    Diasvactomados = Diasvac
    Propor = True


    'Se le descuenta los dias de vacaciones que ya estan marcados como liquidados en el pago /dto de la Gestion integral
    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " INNER JOIN vacpagdesc ON vacpagdesc.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro "
    StrSql = StrSql & " WHERE (empleado = " & buliq_empleado!ternro & " )"
    StrSql = StrSql & " AND (tdnro = 2) "
    StrSql = StrSql & " AND elfechahasta < " & ConvFecha(Fec_Fin)
    StrSql = StrSql & " AND vacpagdesc.pago_dto = 3 and not vacpagdesc.pronro is null "
    OpenRecordset StrSql, rs_Emp_Lic
    
    Do While Not rs_Emp_Lic.EOF
        If rs_Emp_Lic!vacanio = Year(Fec_Fin) Then
            Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
            Propor = True
        Else
            If rs_Emp_Lic!vacanio + 1 = Year(Fec_Fin) Then
                Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
                Propor = False
            End If
        End If
        
        rs_Emp_Lic.MoveNext
    Loop
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Se le descuentan " & (Diasvac - Diasvactomados) & " dias de vacaciones que ya estan marcados como liquidados"
    End If
    

    'PROPORCIONAR  LA CANTIDAD TOTAL DE DIAS CORRESPONDIENTES O LA CANT. PENDIENTE EN FUNCION  A LA FECHA DE BAJA
    If Propor Then
        Diasvac = Diasvactomados / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin))
    Else
        Diasvac = Diasvac / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin)) + Diasvactomados
    End If

    Diasvac = IIf(Fix(Diasvac) = Diasvac, Diasvac, Fix(Diasvac + 1))
    If Diasvac < 0 Then
        Diasvac = 0
    End If

    Valor = IIf(Diasvac < 0, 0, Diasvac)
    Bien = True

End Sub

Public Sub bus_Vac_No_Gozadas_Pendientes(ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Vacaciones no Gozadas
' Autor      : FGZ
' Fecha      : 02/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fec_Fin As Date

Dim Maximo       As Single
Dim Tolerancia   As Single
Dim Inas_Ingreso As Single
Dim Diasvac     As Single
Dim Diasvactomados  As Single
Dim Genera       As Boolean
Dim Propor       As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_DiasVac As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin se toma a la fecha de baja de la ultima fase "
        Flog.writeline Espacios(Tabulador * 4) & "Si no esta dado de baja tomo la fecha de baja prevista y si "
        Flog.writeline Espacios(Tabulador * 4) & " si la fecha de baja prevista es nula tomo la fecha fin del proceso "
    End If
    
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND real = -1 " & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If Not EsNulo(rs_Fases!bajfec) Then
            Fec_Fin = rs_Fases!bajfec
        Else
            If Not EsNulo(buliq_empleado!empfbajaprev) Then
                Fec_Fin = buliq_empleado!empfbajaprev
            Else
                Fec_Fin = buliq_proceso!profecfin
            End If
        End If
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin es:  " & Fec_Fin
    End If

    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Maximo = Arr_Programa(NroProg).Auxint1
        Tolerancia = Arr_Programa(NroProg).Auxint2
        Inas_Ingreso = 0
    Else
        Exit Sub
    End If


'    'A pedido de Analia
'    ' run pronov/des01.p(empleado.ternro, FEC-fin, output diasvac,output genera).
'    'Call bus_DiasVac_Masivos
'    Call bus_DiasVac
'
'    Diasvac = Valor
'    Genera = Bien
'
'    If Not Genera Then
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 4) & "No se pudieron generar los dias de Vacaciones no Gozadas de: " & buliq_empleado!empleg
'            Exit Sub
'        End If
'    End If

    Diasvac = 0
    Genera = Bien
    
    StrSql = "SELECT sum(vacdiascor.vdiascorcant) suma FROM vacdiascor "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
    StrSql = StrSql & " WHERE vacdiascor.ternro =" & buliq_empleado!ternro
    StrSql = StrSql & " AND vacacion.vacanio <= " & Year(Fec_Fin)
    OpenRecordset StrSql, rs_DiasVac
    If Not rs_DiasVac.EOF Then
        Diasvac = IIf(Not EsNulo(rs_DiasVac!Suma), rs_DiasVac!Suma, 0)
    End If

    Diasvactomados = Diasvac
    Propor = True

    'Se le descuenta los dias de vacaciones que ya estan marcados como liquidados en el pago /dto de la Gestion integral
    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro "
    StrSql = StrSql & " WHERE (empleado = " & buliq_empleado!ternro & " )"
    StrSql = StrSql & " AND (tdnro = 2) "
    StrSql = StrSql & " AND elfechahasta < " & ConvFecha(Fec_Fin)
    OpenRecordset StrSql, rs_Emp_Lic
    
    Do While Not rs_Emp_Lic.EOF
        If rs_Emp_Lic!vacanio = Year(Fec_Fin) Then
            Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
            Propor = True
        Else
            If rs_Emp_Lic!vacanio + 1 = Year(Fec_Fin) Then
                Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
                Propor = False
            End If
        End If
        
        rs_Emp_Lic.MoveNext
    Loop

    'PROPORCIONAR  LA CANTIDAD TOTAL DE DIAS CORRESPONDIENTES O LA CANT. PENDIENTE EN FUNCION  A LA FECHA DE BAJA
    If Propor Then
        Diasvac = Diasvactomados / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin))
    Else
        Diasvac = Diasvac / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin)) + Diasvactomados
    End If

    Diasvac = IIf(Fix(Diasvac) = Diasvac, Diasvac, Fix(Diasvac + 1))
    If Diasvac < 0 Then
        Diasvac = 0
    End If

    Valor = IIf(Diasvac < 0, 0, Diasvac)
    Bien = True




End Sub

Public Sub bus_Vac_No_Gozadas_Pendientes_OLD(ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Vacaciones no Gozadas
' Autor      : FGZ
' Fecha      : 02/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fec_Fin As Date

Dim Maximo       As Single
Dim Tolerancia   As Single
Dim Inas_Ingreso As Single
Dim Diasvac     As Single
Dim Diasvactomados  As Single
Dim Genera       As Boolean
Dim Propor       As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_DiasVac As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin se toma a la fecha de baja de la ultima fase "
        Flog.writeline Espacios(Tabulador * 4) & "Si no esta dado de baja tomo la fecha de baja prevista y si "
        Flog.writeline Espacios(Tabulador * 4) & " si la fecha de baja prevista es nula tomo la fecha fin del proceso "
    End If
    
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND real = -1 " & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If Not EsNulo(rs_Fases!bajfec) Then
            Fec_Fin = rs_Fases!bajfec
        Else
            If Not EsNulo(buliq_empleado!empfbajaprev) Then
                Fec_Fin = buliq_empleado!empfbajaprev
            Else
                Fec_Fin = buliq_proceso!profecfin
            End If
        End If
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin es:  " & Fec_Fin
    End If

    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Maximo = Arr_Programa(NroProg).Auxint1
        Tolerancia = Arr_Programa(NroProg).Auxint2
        Inas_Ingreso = 0
    Else
        Exit Sub
    End If


'    'A pedido de Analia
'    ' run pronov/des01.p(empleado.ternro, FEC-fin, output diasvac,output genera).
'    'Call bus_DiasVac_Masivos
'    Call bus_DiasVac
'
'    Diasvac = Valor
'    Genera = Bien
'
'    If Not Genera Then
'        If CBool(USA_DEBUG) Then
'            Flog.writeline Espacios(Tabulador * 4) & "No se pudieron generar los dias de Vacaciones no Gozadas de: " & buliq_empleado!empleg
'            Exit Sub
'        End If
'    End If

    Diasvac = 0
    Genera = Bien
    
    StrSql = "SELECT sum(vacdiascor.vdiascorcant) suma FROM vacdiascor "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
    StrSql = StrSql & " WHERE vacdiascor.ternro =" & buliq_empleado!ternro
    StrSql = StrSql & " AND vacacion.vacanio < " & Year(Fec_Fin)
    OpenRecordset StrSql, rs_DiasVac
    If Not rs_DiasVac.EOF Then
        Diasvac = IIf(Not EsNulo(rs_DiasVac!Suma), rs_DiasVac!Suma, 0)
    End If

    Diasvactomados = Diasvac
    Propor = True

    'Se le descuenta los dias de vacaciones que ya estan marcados como liquidados en el pago /dto de la Gestion integral
    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " INNER JOIN vacpagdesc ON vacpagdesc.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro "
    StrSql = StrSql & " WHERE (empleado = " & buliq_empleado!ternro & " )"
    StrSql = StrSql & " AND (tdnro = 2) "
    StrSql = StrSql & " AND elfechahasta < " & ConvFecha(Fec_Fin)
    StrSql = StrSql & " AND vacpagdesc.pago_dto = 3 and not vacpagdesc.pronro is null "
    OpenRecordset StrSql, rs_Emp_Lic
    
    Do While Not rs_Emp_Lic.EOF
        If rs_Emp_Lic!vacanio = Year(Fec_Fin) Then
            Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
            Propor = True
        Else
            If rs_Emp_Lic!vacanio + 1 = Year(Fec_Fin) Then
                Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
                Propor = False
            End If
        End If
        
        rs_Emp_Lic.MoveNext
    Loop

    Diasvac = Diasvactomados
    If Diasvac < 0 Then
        Diasvac = 0
    End If

    Valor = IIf(Diasvac < 0, 0, Diasvac)
    Bien = True

End Sub



Public Sub Bus_Edad_Empleado()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de la edad del empleado
' Autor      : FGZ
' Fecha      : 05/07/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim OpcionFin As Integer    '0 - Fin del Proceso
                            '1 - Fin del Periodo
                            '2 - Fin de A�o
                            '3 - Fecha Actual
                            
Dim Salida As Integer       '0 - Dias
                            '1 - Meses
                            '2 - A�os
Dim FechaFin As Date

'Dim Param_cur As New ADODB.Recordset
Dim rs_Tercero As New ADODB.Recordset
Dim dias As Integer
Dim Meses As Integer
Dim anios As Integer

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        'Los defaults son A la fecha actual y en A�os
        OpcionFin = IIf(Not EsNulo(Arr_Programa(NroProg).Auxint1), Arr_Programa(NroProg).Auxint1, 3)
        Salida = IIf(Not EsNulo(Arr_Programa(NroProg).Auxint2), Arr_Programa(NroProg).Auxint2, 2)
    Else
        Exit Sub
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametro no configurado: "
        End If
    End If
   
    ' Obtener los parametros de la Busqueda
    StrSql = "SELECT * FROM tercero WHERE ternro = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Tercero
    
    If Not rs_Tercero.EOF Then
        Select Case OpcionFin
        Case 0: 'Fin del Proces
            FechaFin = IIf(Not EsNulo(buliq_proceso!profecfin), buliq_proceso!profecfin, Date)
        Case 1: 'Fin del Periodo
            FechaFin = IIf(Not EsNulo(buliq_periodo!pliqhasta), buliq_periodo!pliqhasta, Date)
        Case 2: 'Fin del A�o Actual
            FechaFin = CDate("31/12/" & Year(Date))
        Case 3: 'Fecha actual
            FechaFin = Date
        End Select
        
        Call DIF_FECHAS2(rs_Tercero!terfecnac, FechaFin, dias, Meses, anios)
        
        Select Case Salida
        Case 0: 'Dias
            Valor = dias + Meses * 360 + anios * 12
            'Valor = DateDiff("d", rs_Tercero!terfecnac, FechaFin)
        Case 1: 'Meses
            Valor = anios * 12 + Meses
            'Valor = DateDiff("m", rs_Tercero!terfecnac, FechaFin)
        Case 2: 'A�os
            Valor = anios
            'Valor = DateDiff("yyyy", rs_Tercero!terfecnac, FechaFin)
        End Select
        Bien = True
    Else
        Bien = False
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se Encuentra el tercero de: " & buliq_empleado!empleg
        End If
        Exit Sub
    End If
    
' Cierro todo y Libero
If rs_Tercero.State = adStateOpen Then rs_Tercero.Close
Set rs_Tercero = Nothing
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

End Sub


Public Sub bus_BaseLicencias()
' ---------------------------------------------------------------------------------------------
' Descripcion: Busca un acumulador una cantidad x de meses a la fecha de inicio de una
'              Licencia de un Tipo X y realizo la operacion sobre los valores encontrados.
' Autor      : FGZ
' Fecha      : 14/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim TipoLicencia As Long    'Tipo de Estructura
Dim Retorna_Cantidad As Boolean     'Retorna cantidad o Monto
                            ' TRUE - Cantidad y FALSE - Monto

Dim CantMeses As Integer        ' Cantidad de meses para atras
Dim Acumulador As Long      'acunro
Dim Opcion As Long          ' 1 - Sumatoria
                            ' 2 - Maximo
                            ' 3 - Promedio
                            ' 4 - Promedio sin 0
                            ' 5 - Minimo
Dim Incluye As Integer      ' 0  - No Incluye
                            ' 1  - Proceso Actual
                            ' 2  - Periodo Actual con Proceso actual
                            ' 3  - Periodo Actual sin proceso actual


Dim FechaDeInicio As Date
Dim FechaDeFin As Date

Dim FechaDesde As Date  'Fecha desde de la primer Licencia del tipo (de ah� tengo que ir para atras x meses)
Dim Continua As Boolean
Dim PliqNro As Long
Dim CantAnios As Integer
Dim UsaActual As Boolean
Dim UsaPeriodoActual As Boolean
Dim Con_Fases As Boolean
Dim Cantidad As Single

Dim MesDesde As Integer
Dim AnioDesde As Integer
Dim MesHasta As Integer
Dim AnioHasta As Integer

'Dim Param_cur As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset
Dim rs_Periodo As New ADODB.Recordset

    ' Busco una Licencia de un tipo particular desde la fecha hasta del proceso
    ' que estoy liquidando hacia atras. Teniendo en cuenta que la licencia encontrada
    ' puede ser continuacion de otra Licencia que termine un dia antes y sea del mismo tipo.
    
    ' Una vez encontrada la fecha inicial de la licencia. Busco un acumulador
    ' una X cantidad de meses para atras realizando alguna operacion con los valores obtenidos (tipicamente promedio).


    Bien = False
    Valor = 0

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur

    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoLicencia = CInt(Arr_Programa(NroProg).Auxchar1)
        Acumulador = CLng(Arr_Programa(NroProg).Auxint1)
        CantMeses = CInt(Arr_Programa(NroProg).Auxint2)
        Opcion = CInt(Arr_Programa(NroProg).Auxint3)
        Incluye = Arr_Programa(NroProg).Auxint4
        Retorna_Cantidad = IIf(Arr_Programa(NroProg).Auxint5 = -1 Or Arr_Programa(NroProg).Auxint5 = 2, False, True)
        Con_Fases = CBool(Arr_Programa(NroProg).Auxlog1)
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros cargados "
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

    'FGZ - 29/01/2004
    FechaDeInicio = buliq_proceso!profecini
    FechaDeFin = buliq_proceso!profecfin

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco una Licencia de un tipo particular desde la fecha hasta del proceso "
        Flog.writeline Espacios(Tabulador * 4) & "que estoy liquidando hacia atras. Teniendo en cuenta que la licencia encontrada "
        Flog.writeline Espacios(Tabulador * 4) & "puede ser continuacion de otra Licencia que termine un dia antes y sea del mismo tipo. "
        Flog.writeline
        Flog.writeline Espacios(Tabulador * 4) & "Una vez encontrada la fecha inicial de la licencia. Busco un acumulador "
        Flog.writeline Espacios(Tabulador * 4) & "una X cantidad de meses para atras realizando alguna operacion con los valores obtenidos (tipicamente promedio). "
    End If

    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " WHERE (empleado = " & buliq_empleado!ternro & " )"
    StrSql = StrSql & " AND tdnro = " & TipoLicencia
    StrSql = StrSql & " AND elfechadesde <=" & ConvFecha(FechaDeFin)
    StrSql = StrSql & " ORDER BY elfechadesde DESC "
    OpenRecordset StrSql, rs_Lic

    If rs_Lic.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� ninguna Licencia de ese tipo " & TipoLicencia
        End If
        Bien = False
        Exit Sub
    Else
        FechaDesde = rs_Lic!elfechadesde
    End If

    rs_Lic.MoveNext
    Continua = True

    Do While Not rs_Lic.EOF And Continua
        If CDate(rs_Lic!elfechahasta + 1) = FechaDesde Then
            FechaDesde = rs_Lic!elfechadesde
        Else
            Continua = False
        End If

        rs_Lic.MoveNext
    Loop

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Fecha desde a buscar : " & FechaDesde
        Flog.writeline Espacios(Tabulador * 4) & "Busco el periodo correspondiente a esa fecha "
    End If

    'Busco el periodo
    StrSql = "SELECT * FROM periodo WHERE pliqmes = " & Month(FechaDesde)
    StrSql = StrSql & " AND pliqanio =" & Year(FechaDesde)
    OpenRecordset StrSql, rs_Periodo
    If rs_Periodo.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontro el Periodo para el mes " & Month(FechaDesde)
        End If
        Exit Sub
    Else
        PliqNro = rs_Periodo!PliqNro
        MesHasta = rs_Periodo!pliqmes
        AnioHasta = rs_Periodo!pliqanio
    End If

    CantAnios = Int(CantMeses / 12)
    CantMeses = CantMeses - (CantAnios * 12)

    Select Case Incluye
    Case 0: 'No icluye ni Periodo actual ni proceso actual
        MesDesde = MesHasta - CantMeses
        AnioDesde = AnioHasta - CantAnios
        If MesHasta = buliq_periodo!pliqmes Then
            If MesHasta = 1 Then
                AnioHasta = AnioHasta - 1
                MesHasta = 12
            Else
                MesHasta = MesHasta - 1
            End If
        End If
        UsaActual = False
        UsaPeriodoActual = False
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No icluye ni Periodo actual ni proceso actual "
            Flog.writeline Espacios(Tabulador * 4) & "Desde el mes " & MesDesde & " del " & AnioDesde & " hasta el mes " & MesHasta & " del " & AnioHasta
        End If
    
    Case 1: ' Incluye Proceso Actual y no periodo actual
        MesDesde = MesHasta - CantMeses
        AnioDesde = AnioHasta - CantAnios
        If MesHasta = buliq_periodo!pliqmes Then
            If MesHasta = 1 Then
                AnioHasta = AnioHasta - 1
                MesHasta = 12
            Else
                MesHasta = MesHasta - 1
            End If
        End If
        UsaActual = True
        UsaPeriodoActual = False
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Incluye Proceso Actual y no periodo actual "
            Flog.writeline Espacios(Tabulador * 4) & "Desde el mes " & MesDesde & " del " & AnioDesde & " hasta el mes " & MesHasta & " del " & AnioHasta
        End If
        
    Case 2: 'Incluye Periodo Actual y el Proceso Actual
        MesDesde = MesHasta - CantMeses
        AnioDesde = AnioHasta - CantAnios
        UsaActual = True
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y el Proceso Actual "
            Flog.writeline Espacios(Tabulador * 4) & "Desde el mes " & MesDesde & " del " & AnioDesde & " hasta el mes " & MesHasta & " del " & AnioHasta
        End If
        
    Case 3: 'Incluye Periodo Actual y no el Proceso Actual
        MesDesde = MesHasta - CantMeses
        AnioDesde = AnioHasta - CantAnios
        UsaActual = False
        UsaPeriodoActual = True
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y no el Proceso Actual "
            Flog.writeline Espacios(Tabulador * 4) & "Desde el mes " & MesDesde & " del " & AnioDesde & " hasta el mes " & MesHasta & " del " & AnioHasta
        End If
    End Select

    Con_Fases = False
    Select Case Opcion
    Case 1: 'Sumatoria
        Call AM_Sum(Acumulador, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Sumatoria "
        End If
        
    Case 2: 'Maximo
        Call AM_Max(Acumulador, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Maximo "
        End If
        
    Case 3: 'Promedio
        Call AM_Prom(Acumulador, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, ((CantAnios * 12) + CantMeses), UsaPeriodoActual)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Promedio "
        End If
        
    Case 4: 'Promedio sin cero
        Call AM_PromSin0(Acumulador, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Promedio sin 0 "
        End If
        
    Case 5: 'Minimo
        Call AM_Min(Acumulador, MesHasta, AnioHasta, CantMeses, CantAnios, Con_Fases, Valor, Cantidad, False, UsaActual, UsaPeriodoActual)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Minimo "
        End If
        
    Case Else
    End Select

    If Retorna_Cantidad Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Retorna cantidad "
        End If
        Valor = Cantidad
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Retorna monto "
        End If
    End If
    Bien = True

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Lic.State = adStateOpen Then rs_Lic.Close
Set rs_Lic = Nothing

If rs_Periodo.State = adStateOpen Then rs_Periodo.Close
Set rs_Periodo = Nothing

End Sub


Public Sub bus_DiasDeIngreso()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo la cantidad de dias desde el inicio en el mes
' Autor      : FGZ
' Fecha      : 05/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim dias As Integer

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim fechasSeteadas As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
    
    Bien = False
    Valor = 0
    
'    ' Obtener los parametros de la Busqueda
'    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        TipoFecha = Arr_Programa(nroprog).auxint1
'    Else
'        Exit Sub
'    End If

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco la ultima fase para el empleado " & buliq_empleado!ternro
    End If

    'Busco la ultima fase
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND real = -1 AND Fases.altfec <= " & ConvFecha(buliq_proceso!profecfin) & _
             " AND (Fases.bajfec >= " & ConvFecha(buliq_proceso!profecini) & " OR Fases.bajfec is null )" & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then
     dias = 0
     Do While Not rs_Fases.EOF

        fechasSeteadas = False
        
        FechaDesde = IIf(rs_Fases!altfec < buliq_proceso!profecini, buliq_proceso!profecini, rs_Fases!altfec)
        
        If Not EsNulo(rs_Fases!bajfec) Then
                FechaHasta = IIf(rs_Fases!bajfec < buliq_proceso!profecfin, rs_Fases!bajfec, buliq_proceso!profecfin)
            Else
                FechaHasta = buliq_proceso!profecfin
            End If
            
        dias = dias + DateDiff("d", FechaDesde, FechaHasta) + 1
        
        'If Not EsNulo(rs_Fases!bajfec) Then
        '   If rs_Fases!bajfec < buliq_proceso!profecini Then
        '      FechaDesde = rs_Fases!bajfec
        '      dias = DateDiff("d", rs_Fases!altfec, rs_Fases!bajfec)
        '      If dias > Day(rs_Fases!bajfec) Then
        '         dias = Day(rs_Fases!bajfec)
        '      End If
        '      fechasSeteadas = True
        '   End If
        'End If
        
        'If Not fechasSeteadas Then
         '   FechaDesde = IIf(rs_Fases!altfec < buliq_proceso!profecini, buliq_proceso!profecini, rs_Fases!altfec)
            
         '   If CBool(USA_DEBUG) Then
         '       Flog.writeline Espacios(Tabulador * 4) & "Fecha desde : " & FechaDesde
         '   End If
            
         '   If CBool(USA_DEBUG) Then
          '      Flog.writeline Espacios(Tabulador * 4) & "si la fecha de de baja es mayor a la fecha de fin del proceso, "
          '      Flog.writeline Espacios(Tabulador * 4) & "la fecha hasta seria la fecha de fin del proceso"
          '  End If
            ' si la fecha de de baja es mayor a la fecha de fin del proceso,
            ' la fecha hasta seria la fecha de fin del proceso
          '  If Not EsNulo(rs_Fases!bajfec) Then
          '      FechaHasta = IIf(rs_Fases!bajfec < buliq_proceso!profecfin, rs_Fases!bajfec, buliq_proceso!profecfin)
          '  Else
          '      FechaHasta = buliq_proceso!profecfin
          '  End If
          '  If CBool(USA_DEBUG) Then
          '      Flog.writeline Espacios(Tabulador * 4) & "Fecha hasta : " & FechaHasta
          '  End If
            
          '  dias = DateDiff("d", FechaDesde, FechaHasta) + 1
            
           If FechaHasta = buliq_proceso!profecfin Then
            
                 Select Case Day(FechaHasta)
                 Case 31: dias = dias - 1
                 Case 28: dias = dias + 2
                 Case 29: dias = dias + 1
                 End Select
            End If
            
        rs_Fases.MoveNext
     Loop
   
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontraron fases para el empleado " & buliq_empleado!ternro
        End If
        Bien = False
        Exit Sub
    End If

    If dias > 30 Then
        dias = 30
    End If
    
    If Month(FechaDesde) = 2 Then
        If Biciesto(Year(FechaDesde)) Then
            If dias = 29 Then
                dias = 30
            End If
        Else
            If dias = 28 Then
                dias = 30
            End If
        End If
    End If
    
    Valor = 30 - dias
    Bien = True

End Sub



Public Sub bus_ValorEnOtroLegajo()
' ---------------------------------------------------------------------------------------------
' Descripcion: Sumo los valores de un concepto o acumulador en el mes que se esta liquidando
'               para otro legajo (si es que existe) que tenga el mismo nro de cuil.
' Autor      : FGZ
' Fecha      : 03/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim EsAcumulador As Boolean
Dim ConAcu As Long
Dim Monto As Boolean

Dim Aux_Monto As Single
Dim Aux_Cantidad As Single
Dim Aux_Cuil As String
Dim Aux_Tercero As Long
Dim Aux_Cliqnro As Long
Dim NoExiste As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Cuil As New ADODB.Recordset
Dim rs_cabliq As New ADODB.Recordset
Dim rs_AcuLiq As New ADODB.Recordset
Dim rs_Detliq As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur

    If Arr_Programa(NroProg).Prognro <> 0 Then
        EsAcumulador = IIf(Arr_Programa(NroProg).Auxint1 = 1, False, True)
        ConAcu = Arr_Programa(NroProg).Auxint3
        Monto = IIf(Arr_Programa(NroProg).Auxint2 = 1, False, True)
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros cargados "
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

    'Busco el nro de cuil del empleado
    StrSql = " SELECT cuil.nrodoc FROM tercero " & _
             " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) " & _
             " WHERE tercero.ternro = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Cuil
    If Not rs_Cuil.EOF Then
        Aux_Cuil = Left(CStr(rs_Cuil!nrodoc), 13)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline "Error al obtener los datos del cuil"
        End If
        Exit Sub
    End If

    'Busco otro empleado de otra empresa que tenga el mismo nro de cuil
    StrSql = " SELECT cuil.nrodoc, tercero.ternro FROM tercero "
    StrSql = StrSql & " INNER JOIN ter_doc cuil ON (tercero.ternro = cuil.ternro AND cuil.tidnro = 10) "
    StrSql = StrSql & " WHERE tercero.ternro <> " & buliq_empleado!ternro
    StrSql = StrSql & " AND nrodoc = '" & Aux_Cuil & "'"
    If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
    OpenRecordset StrSql, rs_Cuil
    
    If rs_Cuil.EOF Then
        If CBool(USA_DEBUG) Then
            Flog.writeline "Error al obtener los datos del cuil"
        End If
        NoExiste = True
    End If
    
    If NoExiste Then
        Valor = 0
        Bien = True
        Exit Sub
    End If
    
    Do While Not rs_Cuil.EOF
        Aux_Tercero = rs_Cuil!ternro
        NoExiste = False

        'busco la cabecera de liquidacion para
        StrSql = "SELECT * FROM proceso "
        StrSql = StrSql & " INNER JOIN cabliq ON cabliq.pronro = proceso.pronro "
        StrSql = StrSql & " INNER JOIN empleado ON cabliq.empleado = empleado.ternro "
        StrSql = StrSql & " WHERE proceso.pliqnro =" & buliq_periodo!PliqNro
        StrSql = StrSql & " AND empleado.ternro =" & Aux_Tercero
        OpenRecordset StrSql, rs_cabliq
        If Not rs_cabliq.EOF Then
            Aux_Cliqnro = rs_cabliq!cliqnro
        Else
            Bien = False
            Aux_Cliqnro = 0
        End If
        If Aux_Cliqnro <> 0 Then
            If EsAcumulador Then
                StrSql = "SELECT alcant, almonto FROM acu_liq "
                StrSql = StrSql & " WHERE cliqnro =" & Aux_Cliqnro
                StrSql = StrSql & " AND acunro =" & ConAcu
                OpenRecordset StrSql, rs_AcuLiq
            
                If Not rs_AcuLiq.EOF Then
                    If Not Monto Then
                        Valor = Valor + rs_AcuLiq!alcant
                    Else
                        Valor = Valor + rs_AcuLiq!almonto
                    End If
                    Bien = True
                Else
                    Bien = True
                    Valor = 0
                End If
            Else
                StrSql = "SELECT dlicant, dlimonto FROM detliq "
                StrSql = StrSql & " WHERE cliqnro =" & Aux_Cliqnro
                StrSql = StrSql & " AND concnro =" & ConAcu
                    OpenRecordset StrSql, rs_Detliq
                If Not rs_Detliq.EOF Then
                    If Not Monto Then
                        Valor = Valor + rs_Detliq!dlicant
                    Else
                        Valor = Valor + rs_Detliq!dlimonto
                    End If
                    Bien = True
                Else
                    Bien = True
                    Valor = 0
                End If
            End If
        End If
        
        rs_Cuil.MoveNext
    Loop
    Bien = True
    
If rs_Cuil.State = adStateOpen Then rs_Cuil.Close
Set rs_Cuil = Nothing

If rs_cabliq.State = adStateOpen Then rs_cabliq.Close
Set rs_cabliq = Nothing
    
If rs_AcuLiq.State = adStateOpen Then rs_AcuLiq.Close
Set rs_AcuLiq = Nothing
    
If rs_Detliq.State = adStateOpen Then rs_Detliq.Close
Set rs_Detliq = Nothing
    
End Sub


Public Sub bus_DiasEnMesSegunFase()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo la cantidad en el mes segun fases.
' Autor      : FGZ
' Fecha      : 12/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim dias As Integer

Dim FechaDesde As Date
Dim FechaHasta As Date

'Dim Param_cur As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
    
    Bien = False
    Valor = 0
    
'    ' Obtener los parametros de la Busqueda
'    StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
'    'OpenRecordset StrSql, Param_cur
'
'    If Arr_Programa(NroProg).Prognro <> 0 Then
'        TipoFecha = Arr_Programa(nroprog).auxint1
'    Else
'        Exit Sub
'    End If

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco la ultima fase para el empleado " & buliq_empleado!ternro
    End If

    'Busco la ultima fase
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND sueldo = -1 " & _
             " AND altfec <= " & ConvFecha(buliq_proceso!profecfin) & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        FechaDesde = IIf(rs_Fases!altfec < buliq_proceso!profecini, buliq_proceso!profecini, rs_Fases!altfec)
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Fecha desde : " & FechaDesde
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "si la fecha de de baja es mayor a la fecha de fin del proceso, "
            Flog.writeline Espacios(Tabulador * 4) & "la fecha hasta seria la fecha de fin del proceso"
        End If
        ' si la fecha de de baja es mayor a la fecha de fin del proceso,
        ' la fecha hasta seria la fecha de fin del proceso
        If Not EsNulo(rs_Fases!bajfec) Then
            If rs_Fases!bajfec >= buliq_proceso!profecini And rs_Fases!bajfec <= buliq_proceso!profecfin Then
                FechaHasta = rs_Fases!bajfec
            Else
                Flog.writeline Espacios(Tabulador * 4) & "No se encontr� fases para el empleado " & buliq_empleado!ternro
                FechaHasta = FechaDesde
            End If
        Else
            'Busco si tiene fecha de baja prevista
            If Not EsNulo(buliq_empleado!empfbajaprev) Then
                If buliq_empleado!empfbajaprev <= buliq_periodo!pliqhasta And buliq_empleado!empfbajaprev >= buliq_periodo!pliqdesde Then
                    FechaHasta = buliq_empleado!empfbajaprev
                Else
                    FechaHasta = buliq_proceso!profecfin
                End If
            Else
                FechaHasta = buliq_proceso!profecfin
            End If
        End If
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Fecha hasta : " & FechaHasta
        End If
        
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontraron fases para el empleado " & buliq_empleado!ternro
        End If
        Bien = False
        Exit Sub
    End If


    dias = DateDiff("d", FechaDesde, FechaHasta) + 1
    If dias > 30 Then
        dias = 30
    End If
    If Month(FechaDesde) = 2 Then
        If Biciesto(Year(FechaDesde)) Then
            If dias = 29 Then
                dias = 30
            End If
        Else
            If dias = 28 Then
                dias = 30
            End If
        End If
    End If
    
    Valor = dias
    Bien = True
End Sub


Public Sub bus_Antiguedad_Por_Acumulador()
' ---------------------------------------------------------------------------------------------
' Descripcion:
'
' Autor      : FGZ
' Fecha      : 25/08/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim NroAcu As Long          ' Nro de Acumulador
Dim Con_Fases As Boolean    ' True  - Calculo con Fases
                            ' False - Calculo sin Fases
Dim Monto As Boolean        ' True  - MOnto
                            ' False - Cantidad
Dim Incluye As Integer      ' 0  - No Incluye
                            ' 1  - Proceso Actual
                            ' 2  - Periodo Actual sin proceso actual
                            ' 3  - Periodo Actual con Proceso actual

Dim Resultado As Integer    ' 1 - en dias
                            ' 2 - en meses
                            ' 3 - en a�os

Dim MesDesde As Integer
Dim AnioDesde As Integer
Dim MesHasta As Integer
Dim AnioHasta As Integer
Dim Cantidad As Single

Dim UsaActual As Boolean

Dim rs_Acu_Mes As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset

    Bien = False
    Valor = 0

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda " & NroProg
    End If
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        NroAcu = Arr_Programa(NroProg).Auxint1
        Resultado = Arr_Programa(NroProg).Auxint3
        Monto = IIf(Arr_Programa(NroProg).Auxint5 = -1 Or Arr_Programa(NroProg).Auxint5 = 2, True, False)
        Incluye = CInt(Arr_Programa(NroProg).Auxint4)
        Con_Fases = CBool(Arr_Programa(NroProg).Auxlog1)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la Busqueda " & NroProg
        End If
        Exit Sub
    End If


Select Case Incluye
Case 0: 'No icluye ni Periodo actual ni proceso actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "No icluye ni Periodo actual ni proceso actual "
    End If

    If buliq_periodo!pliqmes = 1 Then
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
    Else
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
    End If
    UsaActual = False
Case 1: ' Incluye Proceso Actual y no periodo actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Proceso Actual y no periodo actual "
    End If

    If buliq_periodo!pliqmes = 1 Then
        MesHasta = 12
        AnioHasta = buliq_periodo!pliqanio - 1
    Else
        MesHasta = buliq_periodo!pliqmes - 1
        AnioHasta = buliq_periodo!pliqanio
    End If
    UsaActual = True
Case 2: 'Incluye Periodo Actual y el Proceso Actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y el Proceso Actual "
    End If

    MesHasta = buliq_periodo!pliqmes
    AnioHasta = buliq_periodo!pliqanio
    UsaActual = True
Case 3: 'Incluye Periodo Actual y no el Proceso Actual
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Incluye Periodo Actual y no el Proceso Actual "
    End If

    MesHasta = buliq_periodo!pliqmes
    AnioHasta = buliq_periodo!pliqanio
    UsaActual = False
End Select

' Modificado para que tome el promedio para los jornales
If Con_Fases Then
    'Busco la ultima fase activa
    StrSql = "SELECT * FROM fases WHERE estado = -1 AND empleado = " & buliq_empleado!ternro
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
            MesDesde = Month(rs_Fases!altfec)
            AnioDesde = Year(rs_Fases!altfec)
    End If
Else
    MesDesde = Month(buliq_empleado!empfaltagr)
    AnioDesde = Year(buliq_empleado!empfaltagr)
End If

If AnioDesde = AnioHasta Then
    If MesDesde = MesHasta Then
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & NroAcu & _
                 " AND " & AnioDesde & " = amanio " & _
                 " AND ammes =" & MesDesde
    Else
        StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
                 " AND acunro =" & NroAcu & _
                 " AND " & AnioDesde & " = amanio " & _
                 " AND ammes >= " & MesDesde & " AND  ammes <=" & MesHasta
    End If
Else
    StrSql = "SELECT * FROM acu_mes WHERE ternro = " & buliq_empleado!ternro & _
             " AND acunro =" & NroAcu & _
             " AND ((" & AnioDesde & " = amanio AND ammes >= " & MesDesde & ") OR " & _
             " (amanio > " & AnioDesde & " AND amanio < " & AnioHasta & ") OR " & _
             " (ammes <=" & MesHasta & " AND amanio = " & AnioHasta & "))"
End If
StrSql = StrSql & " ORDER BY amanio, ammes"
OpenRecordset StrSql, rs_Acu_Mes

Do While Not rs_Acu_Mes.EOF
    Valor = Valor + IIf(Not IsNull(rs_Acu_Mes!ammonto), rs_Acu_Mes!ammonto, 0)
    Cantidad = Cantidad + IIf(Not IsNull(rs_Acu_Mes!amcant), rs_Acu_Mes!amcant, 0)
    
    rs_Acu_Mes.MoveNext
Loop

' Si es desde el mes actual ==> busco el acu_liq de este proceso
If UsaActual Then
    If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(NroAcu)) Then
        Valor = Valor + objCache_Acu_Liq_Monto.Valor(CStr(NroAcu))
        Cantidad = Cantidad + objCache_Acu_Liq_Cantidad.Valor(CStr(NroAcu))
    End If
End If


If Not Monto Then
    Valor = Cantidad
End If

Select Case Resultado
Case 1: ' En dias
    'dias: REDONDEAR(RESIDUO(D3;30);0)
    Valor = Round((Valor Mod 360) Mod 30, 0)
Case 2: ' En meses
    Valor = Round((Valor Mod 360) Mod 12, 0)
    'meses: TRUNCAR(RESIDUO(D3;360)/30;0)
Case 3: ' en a�os
    'a�os: TRUNCAR(D3/360;0)
    Valor = CInt(Valor / 360)
End Select
Bien = True

' Cierro todo y libero
If rs_Acu_Mes.State = adStateOpen Then rs_Acu_Mes.Close
Set rs_Acu_Mes = Nothing

If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
End Sub


Public Function Dias_Licencias_Mes_Anterior(ByVal Tercero As Long, ByVal FechaDeInicio As Date, ByVal FechaDeFin As Date) As Integer
Dim rs_Lic As New ADODB.Recordset
Dim dias As Integer

    dias = 0

    StrSql = "SELECT * FROM emp_lic WHERE (empleado = " & Tercero & " )" & _
             " AND elfechadesde <=" & ConvFecha(FechaDeFin) & _
             " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
    OpenRecordset StrSql, rs_Lic
        
    Do While Not rs_Lic.EOF
        dias = CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
        rs_Lic.MoveNext
    Loop
        
    If Month(FechaDeInicio) = 2 Then 'Febrero
        If Biciesto(Year(FechaDeInicio)) Then
            If dias >= 29 Then
                dias = 30
            End If
        Else
            If dias >= 28 Then
                dias = 30
            End If
        End If
    Else
        If dias > 30 Then
            dias = 30
        End If
    End If
        
    Dias_Licencias_Mes_Anterior = dias
    
    'cierro
    If rs_Lic.State = adStateOpen Then rs_Lic.Close
    Set rs_Lic = Nothing
End Function

Public Sub bus_Vac_No_Gozadas_A_Pagar(ByVal concnro As Long, ByVal prog As Long)
' ---------------------------------------------------------------------------------------------
' Descripcion: Vacaciones no Gozadas a Pagar
' Autor      : FGZ
' Fecha      : 15/09/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Fec_Fin As Date

Dim Maximo       As Single
Dim Tolerancia   As Single
Dim Inas_Ingreso As Single
Dim Diasvac     As Single
Dim Diasvactomados  As Single
Dim Genera       As Boolean
Dim Propor       As Boolean

'Dim Param_cur As New ADODB.Recordset
Dim rs_Emp_Lic As New ADODB.Recordset
Dim rs_Fases As New ADODB.Recordset
Dim rs_NovEmp As New ADODB.Recordset
Dim rs_DiasVac As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin se toma a la fecha de baja de la ultima fase "
        Flog.writeline Espacios(Tabulador * 4) & "Si no esta dado de baja tomo la fecha de baja prevista y si "
        Flog.writeline Espacios(Tabulador * 4) & " si la fecha de baja prevista es nula tomo la fecha fin del proceso "
    End If
    
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND real = -1 " & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If Not EsNulo(rs_Fases!bajfec) Then
            Fec_Fin = rs_Fases!bajfec
        Else
            If Not EsNulo(buliq_empleado!empfbajaprev) Then
                Fec_Fin = buliq_empleado!empfbajaprev
            Else
                Fec_Fin = buliq_proceso!profecfin
            End If
        End If
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "La fecha fin es:  " & Fec_Fin
    End If

    
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Maximo = Arr_Programa(NroProg).Auxint1
        Tolerancia = Arr_Programa(NroProg).Auxint2
        Inas_Ingreso = 0
    Else
        Exit Sub
    End If

    Diasvac = 0
    Genera = Bien
    
    StrSql = "SELECT sum(vacdiascor.vdiascorcant) suma FROM vacdiascor "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = vacdiascor.vacnro "
    StrSql = StrSql & " WHERE vacdiascor.ternro =" & buliq_empleado!ternro
    StrSql = StrSql & " AND vacacion.vacanio <= " & Year(Fec_Fin)
    OpenRecordset StrSql, rs_DiasVac
    If Not rs_DiasVac.EOF Then
        Diasvac = IIf(Not EsNulo(rs_DiasVac!Suma), rs_DiasVac!Suma, 0)
    End If

    Diasvactomados = Diasvac
    Propor = True


    'Se le descuenta los dias de vacaciones que ya estan marcados como liquidados en el pago /dto de la Gestion integral
    StrSql = "SELECT * FROM emp_lic "
    StrSql = StrSql & " INNER JOIN vacpagdesc ON vacpagdesc.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN lic_vacacion ON lic_vacacion.emp_licnro = emp_lic.emp_licnro "
    StrSql = StrSql & " INNER JOIN vacacion ON vacacion.vacnro = lic_vacacion.vacnro "
    StrSql = StrSql & " WHERE (empleado = " & buliq_empleado!ternro & " )"
    StrSql = StrSql & " AND (tdnro = 2) "
    StrSql = StrSql & " AND elfechahasta < " & ConvFecha(Fec_Fin)
    StrSql = StrSql & " AND vacpagdesc.pago_dto = 3 and not vacpagdesc.pronro is null "
    OpenRecordset StrSql, rs_Emp_Lic
    
    Do While Not rs_Emp_Lic.EOF
        If rs_Emp_Lic!vacanio = Year(Fec_Fin) Then
            Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
            Propor = True
        Else
            If rs_Emp_Lic!vacanio + 1 = Year(Fec_Fin) Then
                Diasvactomados = Diasvactomados - rs_Emp_Lic!elcantdias
                Propor = False
            End If
        End If
        
        rs_Emp_Lic.MoveNext
    Loop
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Se le descuentan " & (Diasvac - Diasvactomados) & " dias de vacaciones que ya estan marcados como liquidados"
    End If
    
    'PROPORCIONAR  LA CANTIDAD TOTAL DE DIAS CORRESPONDIENTES O LA CANT. PENDIENTE EN FUNCION  A LA FECHA DE BAJA
    If Propor Then
        Diasvac = Diasvactomados / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin))
    Else
        Diasvac = Diasvac / 365 * ((Month(Fec_Fin) - 1) * 30 + Day(Fec_Fin)) + Diasvactomados
    End If

    Diasvac = IIf(Fix(Diasvac) = Diasvac, Diasvac, Fix(Diasvac + 1))
    If Diasvac < 0 Then
        Diasvac = 0
    End If

    Valor = IIf(Diasvac < 0, 0, Diasvac)
    Bien = True

End Sub



Public Sub bus_Licencias_Horas_A_Justificar()
' ---------------------------------------------------------------------------------------------
' Descripcion: Retorna las horas a JustficarDias de Licencias entre dos fechas (de un tipo o de todos los tipos)
' Autor      : FGZ
' Fecha      : 14/01/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoLicencia As Long    'Tipo de Estructura
Dim Todas As Boolean        'Todos los tipos

Dim dias As Integer
Dim Horas As Single
Dim Ultimo_Multiplicador_de_Horas As Single
Dim SumaHoras As Integer

Dim SumaDias As Integer
Dim SumaDiasYaGenerados As Integer
Dim FechaDeInicio As Date
Dim FechaDeFin As Date
Dim TipoDia_Ok As Boolean
Dim Dias_Mes_Anterior As Integer

Dim rs_Estructura As New ADODB.Recordset
Dim rs_tipd_con As New ADODB.Recordset
Dim rs_Lic As New ADODB.Recordset


    Bien = False
    Valor = 0

    ' Obtener los parametros de la Busqueda
    If Arr_Programa(NroProg).Prognro <> 0 Then
        Todas = CBool(Arr_Programa(NroProg).Auxlog1)
        If Not Todas Then
            TipoLicencia = Arr_Programa(NroProg).Auxint1
        End If
    Else
        Exit Sub
    End If

'FGZ - 29/01/2004
FechaDeInicio = buliq_proceso!profecini
FechaDeFin = buliq_proceso!profecfin

' Primero Busco  los tipos de dias asociados a los conceptos
If Todas Then 'Todos los tipos de Licencias
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro
Else 'Todos las Licencias del tipo especificado
    StrSql = " SELECT * FROM tipd_con " & _
             " WHERE concnro =" & Buliq_Concepto(Concepto_Actual).concnro & _
             " AND tdnro = " & TipoLicencia
End If
OpenRecordset StrSql, rs_tipd_con

Do While Not rs_tipd_con.EOF
    TipoDia_Ok = True
    If Not EsNulo(rs_tipd_con!tenro) Then
        If rs_tipd_con!tenro <> 0 Then
            StrSql = " SELECT * FROM his_estructura " & _
                     " WHERE ternro = " & buliq_empleado!ternro & " AND " & _
                     " tenro =" & rs_tipd_con!tenro & " AND " & _
                     " estrnro = " & rs_tipd_con!estrnro & " AND " & _
                     " (htetdesde <= " & ConvFecha(FechaDeFin) & ") AND " & _
                     " ((" & ConvFecha(FechaDeFin) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If rs_Estructura.EOF Then
                TipoDia_Ok = False
            End If
        End If
    End If

    If CBool(TipoDia_Ok) Then
        StrSql = "SELECT * FROM emp_lic " & _
                 " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro " & _
                 " WHERE (empleado = " & buliq_empleado!ternro & " )" & _
                 " AND emp_lic.tdnro =" & rs_tipd_con!tdnro & _
                 " AND elfechadesde <=" & ConvFecha(FechaDeFin) & _
                 " AND elfechahasta >= " & ConvFecha(FechaDeInicio)
        OpenRecordset StrSql, rs_Lic

        SumaDias = 0
        SumaHoras = 0
        dias = 0
        Horas = 0
        Do While Not rs_Lic.EOF
            dias = CantidadDeDias(FechaDeInicio, FechaDeFin, rs_Lic!elfechadesde, rs_Lic!elfechahasta)
            
            'reviso si la licencia es completa
            If Todas Then 'Todos los tipos de Licencias
                Dias_Mes_Anterior = Dias_Licencias_Mes_Anterior(buliq_empleado!ternro, DateAdd("m", -1, FechaDeInicio), FechaDeInicio - 1)
                If Dias_Mes_Anterior = 30 Then
                    'calculo los dias reales del mes
                    Dias_Mes_Anterior = DateDiff("d", DateAdd("m", -1, FechaDeInicio), FechaDeInicio - 1) + 1
                    dias = dias + (Dias_Mes_Anterior - 30)
                End If
            Else
                ' solo este tipo
                If rs_Lic!elfechadesde <= DateAdd("m", -1, FechaDeInicio) Then
                    If rs_Lic!elfechahasta >= DateAdd("m", -1, FechaDeFin) Then
                        'Para ajustar la cantidad de dias cuando la lic sobrepasa al mes y fue topeada
                        Dias_Mes_Anterior = DateDiff("d", DateAdd("m", -1, FechaDeInicio), FechaDeInicio - 1) + 1
                        dias = dias + (Dias_Mes_Anterior - 30)
                    End If
                End If
            End If
            'calculo la cantidad de horas
            If Not EsNulo(rs_Lic!tdcanthoras) Then
                Horas = dias * rs_Lic!tdcanthoras
                Ultimo_Multiplicador_de_Horas = rs_Lic!tdcanthoras
            Else
                Horas = 0
                Ultimo_Multiplicador_de_Horas = 0
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Tipo de Licencia sin configuras las Horas. " & rs_Lic!tdnro & " - " & rs_Lic!tddesc
                End If
            End If
            
            SumaDias = SumaDias + dias
            SumaHoras = SumaHoras + Horas
            
            'Marco la licencia para que no se pueda Borrar
            StrSql = "UPDATE emp_lic SET pronro = " & NroProc & _
                     " WHERE emp_licnro = " & rs_Lic!emp_licnro
            objConn.Execute StrSql, , adExecuteNoRecords

            rs_Lic.MoveNext
        Loop
    End If
    rs_tipd_con.MoveNext
Loop

' --------------------------------------------
' SumaDias no debe seperar 30 dias
' --------------------------------------------
'If Month(FechaDeInicio) = 2 Then 'Febrero
'    If Biciesto(Year(FechaDeInicio)) Then
'        If SumaDias >= 29 Then
'            Valor = 30
'        Else
'            Valor = SumaDias
'        End If
'    Else
'        If SumaDias >= 28 Then
'            Valor = 30
'        Else
'            Valor = SumaDias
'        End If
'    End If
'Else
'    If SumaDias > 30 Then
'        Valor = 30
'    Else
'        Valor = SumaDias
'    End If
'End If
Valor = SumaHoras
If Month(FechaDeInicio) = 2 Then 'Febrero
    If Biciesto(Year(FechaDeInicio)) Then
        If SumaDias >= 29 Then
            Valor = SumaHoras - (Ultimo_Multiplicador_de_Horas * (SumaDias - 29))
        Else
            Valor = SumaHoras
        End If
    Else
        If SumaDias >= 28 Then
            Valor = Valor = SumaHoras - (Ultimo_Multiplicador_de_Horas * (SumaDias - 28))
        Else
            Valor = SumaHoras
        End If
    End If
Else
    If SumaDias > 30 Then
        Valor = SumaHoras - (Ultimo_Multiplicador_de_Horas * (SumaDias - 30))
    Else
        Valor = SumaHoras
    End If
End If
Bien = True

' Cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
'Set Param_cur = Nothing

If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
Set rs_Estructura = Nothing

If rs_Lic.State = adStateOpen Then rs_Lic.Close
Set rs_Lic = Nothing

If rs_tipd_con.State = adStateOpen Then rs_tipd_con.Close
Set rs_tipd_con = Nothing

End Sub


Public Sub bus_Feriados_Quincena()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo la cantidad Feriados en la primer quincena o segunda quincena segun fases.
'               y retorna esa cantidad de dias por el monto (tipicamente 9.5)
' Autor      : FGZ
' Fecha      : 06/10/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim dias As Integer
Dim Monto As Single
Dim Quincena As Integer

Dim FechaDesde As Date
Dim FechaHasta As Date
Dim FechaActual As Date
Dim EsFeriado As Boolean

Dim rs_Fases As New ADODB.Recordset
    
Dim objFeriado As New Feriado

    ' inicializacion de variables
    Set objFeriado.Conexion = objConn
   
    Bien = False
    Valor = 0
    dias = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    If Arr_Programa(NroProg).Prognro <> 0 Then
        If IsNumeric(Arr_Programa(NroProg).Auxchar1) Then
            Monto = CSng(Arr_Programa(NroProg).Auxchar1)
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Monto = " & Monto
            End If
        Else
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "El Monto no es numerico " & Monto
                Flog.writeline Espacios(Tabulador * 4) & "Se utilizar� el Monto Default 9.5 "
            End If
            Monto = 9.5
        End If
        Quincena = Arr_Programa(NroProg).Auxint1
        
        If CBool(USA_DEBUG) Then
            If Quincena = 1 Then
                Flog.writeline Espacios(Tabulador * 4) & "Primer Quincena"
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Segunda Quincena"
            End If
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If

    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Busco la ultima fase para el empleado " & buliq_empleado!ternro
    End If
    
    'Busco la ultima fase
    StrSql = "SELECT * FROM fases WHERE empleado = " & buliq_empleado!ternro & _
             " AND sueldo = -1 " & _
             " AND altfec <= " & ConvFecha(buliq_proceso!profecfin) & _
             " ORDER BY altfec "
    OpenRecordset StrSql, rs_Fases
    If Not rs_Fases.EOF Then
        rs_Fases.MoveLast
        If Quincena = 1 Then
            'Primera quincena
            FechaDesde = IIf(rs_Fases!altfec < buliq_proceso!profecini, buliq_proceso!profecini, rs_Fases!altfec)
        Else
            'Segunda quincena
            FechaDesde = IIf(rs_Fases!altfec < CDate("16/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio), CDate("16/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio), rs_Fases!altfec)
        End If
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "si la fecha de de baja es menor a la fecha de fin de la primer quincena, "
            Flog.writeline Espacios(Tabulador * 4) & "la fecha hasta seria el 15 del mes del proceso"
        End If
        If Quincena = 1 Then
            'Primer quincena
            If Not EsNulo(rs_Fases!bajfec) Then
                If rs_Fases!bajfec < CDate("15/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio) Then
                    FechaHasta = rs_Fases!bajfec
                Else
                    FechaHasta = CDate("15/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio)
                End If
            Else
                'Busco si tiene fecha de baja prevista
                If Not EsNulo(buliq_empleado!empfbajaprev) Then
                    If buliq_empleado!empfbajaprev < CDate("15/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio) Then
                        FechaHasta = buliq_empleado!empfbajaprev
                    Else
                        FechaHasta = CDate("15/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio)
                    End If
                Else
                    FechaHasta = CDate("15/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio)
                End If
            End If
        Else
            'Segunda quincena
            If Not EsNulo(rs_Fases!bajfec) Then
                If rs_Fases!bajfec < buliq_proceso!profecfin Then
                    FechaHasta = rs_Fases!bajfec
                Else
                    FechaHasta = buliq_proceso!profecfin
                End If
            Else
                'Busco si tiene fecha de baja prevista
                If Not EsNulo(buliq_empleado!empfbajaprev) Then
                    If buliq_empleado!empfbajaprev < buliq_proceso!profecfin Then
                        FechaHasta = buliq_empleado!empfbajaprev
                    Else
                        FechaHasta = buliq_proceso!profecfin
                    End If
                Else
                    FechaHasta = buliq_proceso!profecfin
                End If
            End If
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� fases para el empleado " & buliq_empleado!ternro
        End If
        FechaDesde = buliq_proceso!profecini
        FechaHasta = CDate("15/" & buliq_periodo!pliqmes & "/" & buliq_periodo!pliqanio)
    End If
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Fecha desde : " & FechaDesde
    End If
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Fecha hasta : " & FechaHasta
    End If

    FechaActual = FechaDesde
    Do While FechaActual <= FechaHasta
        
        EsFeriado = objFeriado.Feriado(FechaActual, buliq_empleado!ternro, False)
        
        If EsFeriado Then
            dias = dias + 1
        End If
        FechaActual = FechaActual + 1
    Loop
    
    Valor = dias * Monto
    Bien = False
    
'Cierro todo
If rs_Fases.State = adStateOpen Then rs_Fases.Close
Set rs_Fases = Nothing
    
End Sub


Public Sub bus_DiasVac_Antig2()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion del valor de la escala para vacaciones pero la antiguedad.
'               se calcula a la fecha de alta del empleado (empleado.empfaltagr) y no con fases.
' Autor      : FGZ
' Fecha      :
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Valor_Grilla(10) As Boolean ' Elemento de una coordenada de una grilla
Dim tipoBus As Long
Dim concnro As Long
Dim prog As Long

Dim tdinteger3 As Integer

Dim ValAnt As Single
Dim Busq As Integer

Dim j As Integer
Dim antig As Boolean
Dim pvariable As Boolean
Dim pvar As Integer
Dim ant As Integer
Dim Continuar As Boolean
Dim Parametros(5) As Integer
Dim grilla_val(10) As Boolean     ' para alojar los valores de:  valgrilla.val(i)

Dim vgrcoor_ant As Long
Dim vgrcoor_pvar As Long

Dim rs_valgrilla As New ADODB.Recordset
Dim rs_cabgrilla As New ADODB.Recordset
Dim rs_tbase As New ADODB.Recordset
Dim TipoBase As Long

Dim NroBusqueda As Long

Dim antdia As Integer
Dim antmes As Integer
Dim antanio As Integer
Dim q As Integer

Dim ValorCoord As Single
Dim Encontro As Boolean
Dim ternro As Long

Dim DiasProporcion As Integer
Dim FactorDivision As Integer

Dim NroVac As Long
Dim cantdias As Integer
Dim Columna As Integer
Dim NroGrilla As Long
'Dim Param_cur As New ADODB.Recordset
Dim dias_trabajados As Integer

Dim FechaAux As Date
Dim Grilla_Ok As Boolean

    Bien = False
    Valor = 0
    ternro = buliq_empleado!ternro
    
    Call Politica(1501, Empleado_Fecha_Fin, Grilla_Ok)
    If Not Grilla_Ok Then
        Flog.writeline "Error cargando configuracion de la Politica 1501"
        Exit Sub
    Else
        DiasProporcion = st_CantidadDias
        FactorDivision = st_FactorDivision
        If FactorDivision = 0 Then
            FactorDivision = 1
        End If
    End If
    
    Call Politica(1502, Empleado_Fecha_Fin, Grilla_Ok)
    If Not Grilla_Ok Then
        Flog.writeline "Error cargando configuracion de la Politica 1502"
        Exit Sub
    Else
        NroGrilla = st_Escala
    End If

    StrSql = "SELECT * FROM cabgrilla " & _
             " WHERE cabgrilla.cgrnro = " & NroGrilla
    OpenRecordset StrSql, rs_cabgrilla

    If rs_cabgrilla.EOF Then
        'La escala de Vacaciones no esta configurada en el tipo de dia para vacaciones
        Exit Sub
    End If
    
    'El tipo Base de la antiguedad
    TipoBase = 4
    
    Continuar = True
    ant = 1
    Do While (ant <= rs_cabgrilla!cgrdimension) And Continuar
        Select Case ant
        Case 1:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_1
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
            
        Case 2:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_2
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 3:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_3
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 4:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_4
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        Case 5:
            StrSql = "SELECT tipoprog.tprogbase FROM programa " & _
                     " INNER JOIN tipoprog ON programa.tprognro = tipoprog.tprognro" & _
                     " WHERE programa.prognro = " & rs_cabgrilla!grparnro_5
            OpenRecordset StrSql, rs_tbase
        
            If Not rs_tbase.EOF Then
                If TipoBase = rs_tbase!tprogbase Then
                    Continuar = False
                Else
                    ant = ant + 1
                End If
            End If
        End Select
    Loop


    For j = 1 To rs_cabgrilla!cgrdimension
        Select Case j
        Case ant:
            'la busqueda es la de antiguedad
            Select Case ant
            Case 1:
                NroProg = rs_cabgrilla!grparnro_1
                'Call bus_Anti0(antdia, antmes, antanio)
                Call bus_Antiguedad_A_FechaAlta(antdia, antmes, antanio)
            Case 2:
                NroProg = rs_cabgrilla!grparnro_2
                'Call bus_Anti0(antdia, antmes, antanio)
                Call bus_Antiguedad_A_FechaAlta(antdia, antmes, antanio)
            Case 3:
                NroProg = rs_cabgrilla!grparnro_3
                'Call bus_Anti0(antdia, antmes, antanio)
                Call bus_Antiguedad_A_FechaAlta(antdia, antmes, antanio)
            Case 4:
                NroProg = rs_cabgrilla!grparnro_4
                'Call bus_Anti0(antdia, antmes, antanio)
                Call bus_Antiguedad_A_FechaAlta(antdia, antmes, antanio)
            Case 5:
                NroProg = rs_cabgrilla!grparnro_5
                'Call bus_Anti0(antdia, antmes, antanio)
                Call bus_Antiguedad_A_FechaAlta(antdia, antmes, antanio)
            End Select
            
            'OJO - Supuestamente este tipo de busqueda esta retornando el resultado en meses
            ' si esta busqueda no retorna meses, no va a encontrar el valor
            Parametros(j) = Valor 'Valor trae cantidad de meses
            
'            Call bus_Antiguedad("VACACIONES", CDate("31/12/" & buliq_periodo!pliqanio), antdia, antmes, antanio, q)
'            Parametros(j) = (antanio * 12) + antmes
        Case Else:
            Select Case j
            Case 1:
                NroProg = rs_cabgrilla!grparnro_1
                Call bus_Estructura
            Case 2:
                NroProg = rs_cabgrilla!grparnro_2
                Call bus_Estructura
            Case 3:
                NroProg = rs_cabgrilla!grparnro_3
                Call bus_Estructura
            Case 4:
                NroProg = rs_cabgrilla!grparnro_4
                Call bus_Estructura
            Case 5:
                NroProg = rs_cabgrilla!grparnro_5
                Call bus_Estructura
            End Select
            Parametros(j) = Valor
        End Select
    Next j

    'Busco la primera antiguedad de la escala menor a la del empleado
    ' de abajo hacia arriba
    StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
    For j = 1 To rs_cabgrilla!cgrdimension
        If j <> ant Then
            StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
        End If
    Next j
        StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
    OpenRecordset StrSql, rs_valgrilla


    Encontro = False
    Do While Not rs_valgrilla.EOF And Not Encontro
'        Call CargarValoresdelaGrilla(rs_valgrilla, grilla_val)
'
'        If ValorCoord >= grilla_val(ant) Then
'            Call BusValor(7, Valor_Grilla, grilla_val, valor, Columna)
'            Bien = True
'        End If
'
'        rs_valgrilla.MoveNext
        Select Case ant
        Case 1:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_1 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 2:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_2 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 3:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_3 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 4:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_4 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        Case 5:
            If Parametros(ant) >= rs_valgrilla!vgrcoor_5 Then
                 If rs_valgrilla!vgrvalor <> 0 Then
                    cantdias = rs_valgrilla!vgrvalor
                    Encontro = True
                    Columna = rs_valgrilla!vgrorden
                 End If
            End If
        End Select
                    
        rs_valgrilla.MoveNext
    Loop
    
    If Not Encontro Then
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No encontr� valor en la escala "
        End If
    
        'Busco si existe algun valor para la estructura y ...
        'si hay que carga la columna correspondiente
        StrSql = "SELECT * FROM valgrilla WHERE cgrnro = " & NroGrilla
        StrSql = StrSql & " AND vgrvalor is not null"
        For j = 1 To rs_cabgrilla!cgrdimension
            If j <> ant Then
                StrSql = StrSql & " AND vgrcoor_" & j & "= " & Parametros(j)
            End If
        Next j
        'StrSql = StrSql & " ORDER BY vgrcoor_" & ant & " DESC "
        OpenRecordset StrSql, rs_valgrilla
        If Not rs_valgrilla.EOF Then
            Columna = rs_valgrilla!vgrorden
        Else
            Columna = 1
        End If
        
        
        dias_trabajados = ((antanio * 365) + (antmes * 30) + antdia)
    
        If Parametros(ant) <= 6 Then

            'FactorDivision = 1
            If DiasProporcion = 20 Then
                If (dias_trabajados / DiasProporcion) / 7 * 5 > Fix((dias_trabajados / DiasProporcion) / 7 * 5) Then
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5) + 1
                Else
                    cantdias = Fix((dias_trabajados / DiasProporcion) / 7 * 5)
                End If
            Else
                cantdias = Round((dias_trabajados / DiasProporcion) / FactorDivision, 0)
            End If
            Valor = cantdias
            Bien = True
            
        Else
            Flog.writeline "No se encontro la escala para el convenio"
            Bien = False
        End If
    Else
        Valor = cantdias
        Bien = True
    End If
   
    
' Cierro todo y libero
If rs_cabgrilla.State = adStateOpen Then rs_cabgrilla.Close
If rs_valgrilla.State = adStateOpen Then rs_valgrilla.Close

Set rs_cabgrilla = Nothing
Set rs_valgrilla = Nothing
End Sub


Public Sub bus_MesPagoProceso()
' ---------------------------------------------------------------------------------------------
' Descripcion: Mes de la fecha de pago de proceso de liquidacion
' Autor      : FGZ
' Fecha      : 15/11/2004
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
    Bien = False
    Valor = 0
    
    Valor = Month(buliq_proceso!profecpago)
    Bien = True
    
End Sub

Public Sub bus_Partes_Diarios()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Partes Diarios
' Autor      : Scarpa D.
' Fecha      : 11/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Actual As Boolean    'Si es del periodo actual o del anterior
                            
Dim TipoParte As Integer 'Tipo de parte a buscar
                            
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Mes As Integer
Dim Anio As Integer
Dim rsConsulta As New ADODB.Recordset
Dim StrSql1 As String
Dim total As Single

    'Obtengo los parametros de la busqueda
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoParte = Arr_Programa(NroProg).Auxint1
        Actual = Arr_Programa(NroProg).Auxlog1
        
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Buscando partes de tipo " & TipoParte
            If Actual Then
                Flog.writeline Espacios(Tabulador * 4) & "Periodo Actual"
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Periodo Anterior"
            End If
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If
    
    'Busco el rango de fechas en el cual tengo que buscar los partes
    If Actual Then
       FechaDesde = buliq_periodo!pliqdesde
       FechaHasta = buliq_periodo!pliqhasta
    Else
       If CInt(buliq_periodo!pliqmes) = 1 Then
          Anio = CInt(buliq_periodo!pliqanio) - 1
          Mes = 12
       Else
          Anio = CInt(buliq_periodo!pliqanio)
          Mes = CInt(buliq_periodo!pliqmes) - 1
       End If
       
       StrSql1 = " SELECT * FROM periodo WHERE pliqmes = " & Mes & " AND pliqanio = " & Anio
       
       OpenRecordset StrSql1, rsConsulta
       
       If rsConsulta.EOF Then
          If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encuentra el periodo para el mes " & Mes & " y anio " & Anio
          End If
          Exit Sub
       Else
          FechaDesde = rsConsulta!pliqdesde
          FechaHasta = rsConsulta!pliqhasta
       End If
       
       rsConsulta.Close
       
    End If
    
    'Busco los partes en el rango de fechas
    StrSql1 = " SELECT sum(observ63) AS Total "
    StrSql1 = StrSql1 & " From ee_parte "
    StrSql1 = StrSql1 & " Where ternro = " & buliq_empleado!ternro
    StrSql1 = StrSql1 & " AND fecnov63 >= " & ConvFecha(FechaDesde)
    StrSql1 = StrSql1 & " AND fecnov63 <= " & ConvFecha(FechaHasta)
    StrSql1 = StrSql1 & " AND tparnro = " & TipoParte
       
    OpenRecordset StrSql1, rsConsulta
    
    If rsConsulta.EOF Then
       total = 0
    Else
       If IsNull(rsConsulta!total) Then
          total = 0
       Else
          total = CSng(rsConsulta!total)
       End If
    End If
    
    rsConsulta.Close

Bien = True
Valor = total
End Sub

Public Sub bus_BAE()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de BAE
' Autor      : Scarpa D.
' Fecha      : 11/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Actual As Boolean    'Si es del periodo actual o del anterior
                            
Dim TipoBAE As Integer 'Tipo de BAE
                            
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim Mes As Integer
Dim Anio As Integer
Dim rsConsulta As New ADODB.Recordset
Dim StrSql1 As String
Dim total As Single

    'Obtengo los parametros de la busqueda
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoBAE = Arr_Programa(NroProg).Auxint1
        Actual = Arr_Programa(NroProg).Auxlog1
        
        If CBool(USA_DEBUG) Then
            If Actual Then
                Flog.writeline Espacios(Tabulador * 4) & "Periodo Actual"
            Else
                Flog.writeline Espacios(Tabulador * 4) & "Periodo Anterior"
            End If
        End If
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If
    
    'Busco el rango de fechas en el cual tengo que buscar
    If Actual Then
       FechaDesde = buliq_periodo!pliqdesde
       FechaHasta = buliq_periodo!pliqhasta
    Else
       If CInt(buliq_periodo!pliqmes) = 1 Then
          Anio = CInt(buliq_periodo!pliqanio) - 1
          Mes = 12
       Else
          Anio = CInt(buliq_periodo!pliqanio)
          Mes = CInt(buliq_periodo!pliqmes) - 1
       End If
       
       StrSql1 = " SELECT * FROM periodo WHERE pliqmes = " & Mes & " AND pliqanio = " & Anio
       
       OpenRecordset StrSql1, rsConsulta
       
       If rsConsulta.EOF Then
          If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encuentra el periodo para el mes " & Mes & " y anio " & Anio
          End If
          Exit Sub
       Else
          FechaDesde = rsConsulta!pliqdesde
          FechaHasta = rsConsulta!pliqhasta
       End If
       
       rsConsulta.Close
       
    End If
    
    'Busco los partes en el rango de fechas
    Select Case TipoBAE
      Case 1 'Cuota
          StrSql1 = " SELECT sum(baecuota) AS Total "
      Case 2 'Bonificacion Adicional
          StrSql1 = " SELECT sum(bonadicional) AS Total "
      Case 3 'Bonificacion Basica
          StrSql1 = " SELECT sum(bonbasica) AS Total "
      Case Else
          StrSql1 = " SELECT sum(baecuota) AS Total "
    End Select
    
    StrSql1 = StrSql1 & " From ee_bae "
    StrSql1 = StrSql1 & " Where ternro = " & buliq_empleado!ternro
    StrSql1 = StrSql1 & " AND aaaamm70 >= " & ConvFecha(FechaDesde)
    StrSql1 = StrSql1 & " AND aaaamm70 <= " & ConvFecha(FechaHasta)
       
    OpenRecordset StrSql1, rsConsulta
    
    If rsConsulta.EOF Then
       total = 0
    Else
       If IsNull(rsConsulta!total) Then
          total = 0
       Else
          total = CSng(rsConsulta!total)
       End If
    End If
    
    rsConsulta.Close

Bien = True
Valor = total

End Sub


Public Sub bus_Movilidad()
' ---------------------------------------------------------------------------------------------
' Descripcion: Calculo de Movilidad
' Autor      : Scarpa D.
' Fecha      : 11/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim tipo As Integer 'Tipo de movilidad
                            
Dim FechaDesde As Date
Dim FechaHasta As Date
Dim rsConsulta As New ADODB.Recordset
Dim StrSql1 As String
Dim total As Single

    'Obtengo los parametros de la busqueda
    If Arr_Programa(NroProg).Prognro <> 0 Then
        tipo = Arr_Programa(NroProg).Auxint1
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busqueda no encontrada " & NroProg
        End If
        Exit Sub
    End If
    
    'Busco los partes en el rango de fechas
    Select Case tipo
      Case 1 'Total
          StrSql1 = " SELECT sum(imptot73) AS Total "
      Case 2 'Liq para ganancias
          StrSql1 = " SELECT sum(impliq73) AS Total "
      Case Else
          StrSql1 = " SELECT sum(imptot73) AS Total "
    End Select
    
    StrSql1 = StrSql1 & " From ee_movilidad "
    StrSql1 = StrSql1 & " Where ternro = " & buliq_empleado!ternro
    StrSql1 = StrSql1 & " AND fecha73 >= " & ConvFecha(buliq_proceso!profecini)
    StrSql1 = StrSql1 & " AND fecha73 <= " & ConvFecha(buliq_proceso!profecfin)
       
    OpenRecordset StrSql1, rsConsulta
    
    If rsConsulta.EOF Then
       total = 0
    Else
       If IsNull(rsConsulta!total) Then
          total = 0
       Else
          total = CSng(rsConsulta!total)
       End If
    End If
    
    rsConsulta.Close

Bien = True
Valor = total

End Sub


Public Sub bus_Cant_Empl_Estr()
' ---------------------------------------------------------------------------------------------
' Descripcion: Cantidad de empleados en la estructura
' Autor      : Scarpa D.
' Fecha      : 23/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------

Dim TipoEstr As Long      ' Tipo de Estructura
Dim ListaEstr As String   ' Lista de estructuras a considerar
Dim StrSql1 As String
                            
'Dim Param_cur As New ADODB.Recordset
Dim rs_Estr As New ADODB.Recordset

    Bien = False
    Valor = 0
   
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        TipoEstr = Arr_Programa(NroProg).Auxint1
        ListaEstr = Arr_Programa(NroProg).Auxchar
    Else
        Exit Sub
    End If
    
    StrSql1 = " SELECT DISTINCT count(his_estructura.ternro) AS Cantidad "
    StrSql1 = StrSql1 & " From his_estructura"
    StrSql1 = StrSql1 & " INNER JOIN empresa ON empresa.estrnro = his_estructura.estrnro"
    StrSql1 = StrSql1 & " INNER JOIN his_estructura hisestr1 ON hisestr1.ternro = his_estructura.ternro AND hisestr1.tenro=" & TipoEstr
    StrSql1 = StrSql1 & " WHERE his_estructura.htetdesde <= " & ConvFecha(buliq_proceso!profecfin)
    StrSql1 = StrSql1 & "   AND (his_estructura.htethasta >= " & ConvFecha(buliq_proceso!profecfin) & " OR his_estructura.htethasta IS NULL)"
    StrSql1 = StrSql1 & "   AND his_estructura.tenro  = 10 "
    StrSql1 = StrSql1 & "   AND empresa.empnro = " & NroEmp
    StrSql1 = StrSql1 & "   AND hisestr1.htetdesde <= " & ConvFecha(buliq_proceso!profecfin)
    StrSql1 = StrSql1 & "   AND (hisestr1.htethasta >= " & ConvFecha(buliq_proceso!profecfin) & " OR hisestr1.htethasta IS NULL)"
    StrSql1 = StrSql1 & "   AND hisestr1.estrnro IN (" & ListaEstr & ")"
    StrSql1 = StrSql1 & "   AND hisestr1.tenro=" & TipoEstr

    OpenRecordset StrSql1, rs_Estr
    
    If rs_Estr.EOF Then
       Valor = 0
    Else
       If IsNull(rs_Estr!Cantidad) Then
          Valor = 0
       Else
          Valor = CLng(rs_Estr!Cantidad)
       End If
    End If

    Bien = True
    
' Cierro todo y libero
If rs_Estr.State = adStateOpen Then rs_Estr.Close
Set rs_Estr = Nothing

End Sub


Public Sub bus_Embargos()
' ---------------------------------------------------------------------------------------------
' Descripcion: Obtencion de los embargos de cualquier tipo de prestamos.
'              Mensuales, 1era o 2da quincena.
'
' Autor      : GdeCos
' Fecha      : 25/02/2005
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim Final As Boolean    'Liquidacion Final
Dim Cancela As Boolean  'Si cancela las cuotas de embargos o no
Dim Nrotpe As Long       'Tipo de Embargos
Dim Opcion As Integer   ' 1 - Mensual
                        ' 2 - 1er Quincena
                        ' 3 - 2da Quincena
Dim Monto As Long       'Monto del Acumulador del Embargos

Dim rs_Embargo As New ADODB.Recordset
Dim rs_Cuota As New ADODB.Recordset
Dim rs_Aux_Cuota As New ADODB.Recordset
Dim rs_TipoEmbargo As New ADODB.Recordset
Dim rs_Acums As New ADODB.Recordset

    Bien = False
    Valor = 0
    
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "Obtener los parametros de la Busqueda "
    End If
    ' Obtener los parametros de la Busqueda
    'StrSql = "SELECT * FROM programa WHERE prognro = " & CStr(NroProg)
    'OpenRecordset StrSql, Param_cur
    
    If Arr_Programa(NroProg).Prognro <> 0 Then
        If EsNulo(Arr_Programa(NroProg).Auxint1) Then
            Nrotpe = -1 ' Todos
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Todos los Tipos de Embargos "
            End If
        Else
            If Arr_Programa(NroProg).Auxint1 = 0 Then
                Nrotpe = -1 ' Todos
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Todos los Tipos de Embargos "
                End If
            Else
                Nrotpe = Arr_Programa(NroProg).Auxint1
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Tipo de Embargo: " & Nrotpe
                End If
            End If
        End If
        
        Opcion = Arr_Programa(NroProg).Auxint2
        
        Final = CBool(Arr_Programa(NroProg).Auxlog1)
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Parametros cargados "
        End If
        
        Cancela = IIf(Not EsNulo(Arr_Programa(NroProg).Auxlog2), CBool(Arr_Programa(NroProg).Auxlog2), True)
    Else
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "No se encontr� la busqueda " & NroProg
        End If
        Exit Sub
    End If


If Final Then 'se trata de una liq. final se descuentan todas las cuotas
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "se trata de una liq. final se descuentan todas las cuotas "
    End If

    Select Case Opcion
    Case 1: 'Embargos Mensuales
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los embargos mensuales "
        End If
        StrSql = "SELECT * FROM embargo "
        StrSql = StrSql & " WHERE embargo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  embargo.embest = 'E' "
        StrSql = StrSql & " AND embargo.embquincenal = 0 "
        If Nrotpe <> -1 Then
            StrSql = StrSql & " AND embargo.tpenro =" & Nrotpe
        End If
        OpenRecordset StrSql, rs_Embargo
        
        If CBool(USA_DEBUG) Then
            If rs_Embargo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron embargos"
            End If
        End If
        Do While Not rs_Embargo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el embargo " & rs_Embargo!embnro
            End If
            StrSql = "SELECT * FROM embcuota "
            StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
            StrSql = StrSql & " AND embcuota.embccancela = 0 "
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!embcimp
                
                If Cancela Then
                    StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", embccancela = -1 "
                    StrSql = StrSql & ", embcimpreal = " & rs_Cuota!embcimp
                    StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                    StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                Bien = True
                rs_Cuota.MoveNext
            Loop
            
            If Cancela Then
                StrSql = "UPDATE embargo SET embest = 'F'"
                StrSql = StrSql & " WHERE embnro = " & rs_Embargo!embnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_Embargo.MoveNext
        Loop
        
    Case 2, 3: 'Embargos de la Primera y Segunda Quincena
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los embargos de la primera y segunda quincena "
        End If
    
        StrSql = "SELECT * FROM embargo "
        StrSql = StrSql & " WHERE embargo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  embargo.embest = 'E' "
        StrSql = StrSql & " AND embargo.embquincenal = -1 "
        If Nrotpe <> -1 Then
            StrSql = StrSql & " AND embargo.tpenro =" & Nrotpe
        End If
        OpenRecordset StrSql, rs_Embargo
        
        If CBool(USA_DEBUG) Then
            If rs_Embargo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron embargos "
            End If
        End If
        
        Do While Not rs_Embargo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el embargo " & rs_Embargo!prenro
            End If
            StrSql = "SELECT * FROM embcuota "
            StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
            StrSql = StrSql & " AND embcuota.embccancela = 0 "
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!embcimp
                
                If Cancela Then
                    StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", embccancela = -1 "
                    StrSql = StrSql & ", embcimpreal = " & rs_Cuota!embcimp
                    StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                    StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
            
            If Cancela Then
                StrSql = "UPDATE embargo SET embest = 'F'"
                StrSql = StrSql & " WHERE embnro = " & rs_Embargo!embnro
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
            
            rs_Embargo.MoveNext
        Loop
    End Select
Else 'liquidacion mensual o quincenal
    If CBool(USA_DEBUG) Then
        Flog.writeline Espacios(Tabulador * 4) & "se trata de una liquidacion mensual o quincenal "
    End If

    Select Case Opcion
    Case 1: 'Embargos Mensuales
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los prestamos mensuales "
        End If
        StrSql = "SELECT * FROM embargo "
        StrSql = StrSql & " WHERE embargo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  embargo.embest = 'E' "
        StrSql = StrSql & " AND embargo.embquincenal = 0 "
        If Nrotpe <> -1 Then
            StrSql = StrSql & " AND embargo.tpenro =" & Nrotpe
        End If
        OpenRecordset StrSql, rs_Embargo
        
        If CBool(USA_DEBUG) Then
            If rs_Embargo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron embargos "
            End If
        End If
        
        If Not rs_Embargo.EOF Then
            StrSql = "SELECT embargo.tpenro, tipoemb.tpefordesc FROM embargo, tipoemb "
            StrSql = StrSql & " WHERE embargo.embnro = " & rs_Embargo!embnro
            StrSql = StrSql & " AND embargo.tpenro = tipoemb.tpenro"
            OpenRecordset StrSql, rs_TipoEmbargo
        End If
        
        If CBool(USA_DEBUG) Then
            If rs_TipoEmbargo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontro el tipo de embargo "
            End If
        End If
        
        Do While Not rs_Embargo.EOF
            
            If (rs_TipoEmbargo!tpefordesc < 1) Then
            
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el embargo " & rs_Embargo!embnro
                End If
            
                StrSql = "SELECT * FROM embcuota "
                StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
                StrSql = StrSql & " AND embcuota.embccancela = 0 "
                StrSql = StrSql & " AND embcuota.embcanio = " & buliq_periodo!pliqanio
                StrSql = StrSql & " AND embcuota.embcmes = " & buliq_periodo!pliqmes
                OpenRecordset StrSql, rs_Cuota
                
                If CBool(USA_DEBUG) Then
                    If rs_Cuota.EOF Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                    End If
                End If
                
                Do While Not rs_Cuota.EOF
                    Valor = Valor + rs_Cuota!embcimp
                    
                    If Cancela Then
                        StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & ", embccancela = -1 "
                        StrSql = StrSql & ", embcimpreal = " & rs_Cuota!embcimp
                        StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                        StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    Bien = True
                    
                    rs_Cuota.MoveNext
                Loop
            Else
                If CBool(USA_DEBUG) Then
                    Flog.writeline Espacios(Tabulador * 4) & "Busco los aumuladores para el embargo " & rs_Embargo!embnro
                End If
            
                StrSql = "SELECT * FROM embcuota "
                StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
                StrSql = StrSql & " AND embcuota.embccancela = 0 "
                StrSql = StrSql & " AND embcuota.embcanio = " & buliq_periodo!pliqanio
                StrSql = StrSql & " AND embcuota.embcmes = " & buliq_periodo!pliqmes
                OpenRecordset StrSql, rs_Cuota
                
                If rs_Cuota.EOF Then
                    If CBool(USA_DEBUG) Then
                        Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas para este periodo"
                    End If
                Else
                    StrSql = "SELECT emb_acu.embacporc, emb_acu.acnro FROM emb_acu"
                    StrSql = StrSql & " WHERE emb_acu.embnro =" & rs_Embargo!embnro
                    OpenRecordset StrSql, rs_Acums
                               
                    If CBool(USA_DEBUG) Then
                        If rs_Acums.EOF Then
                            Flog.writeline Espacios(Tabulador * 4) & "No se encontraron acumuladores "
                        End If
                    End If
                
    
                    Do While Not rs_Acums.EOF
                    'toDo: Falta controlar que pasa si no se puede descontar todo el porcentaje
                        If objCache_Acu_Liq_Monto.EsSimboloDefinido(CStr(rs_Acums!acnro)) Then
                            Monto = objCache_Acu_Liq_Monto.Valor(CStr(rs_Acums!acnro))
                        End If
                        
                        Valor = Valor + ((Monto * rs_Acums!embacporc) / 100)
                        
                        rs_Acums.MoveNext
                    Loop
                        
                    If Cancela Then
                        StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & ", embccancela = -1 "
                        StrSql = StrSql & ", embcimp = " & Valor
                     ' Se deberia controlar cuanto se pudo descontar y no asignar valor directamente
                        StrSql = StrSql & ", embcimpreal = " & Valor
                        StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                        StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                        StrSql = StrSql & ", embcimp = " & Valor
                        StrSql = StrSql & ", embcimpreal = " & Valor
                        StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                        StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    
                    Bien = True
                    
                End If
                
            End If
                                                   
            rs_Embargo.MoveNext
         Loop
                                                  
        
    Case 2: 'Embargos de la Primera quincena
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los embargos de la primera quincena "
        End If
    
        StrSql = "SELECT * FROM embargo "
        StrSql = StrSql & " WHERE embargo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  embargo.embest = 'E' "
        StrSql = StrSql & " AND embargo.embquincenal = -1 "
        If Nrotpe <> -1 Then
            StrSql = StrSql & " AND embargo.tpenro =" & Nrotpe
        End If
        OpenRecordset StrSql, rs_Embargo
        
        If CBool(USA_DEBUG) Then
            If rs_Embargo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron embargos "
            End If
        End If
        
        Do While Not rs_Embargo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el embargo " & rs_Embargo!prenro
            End If
            StrSql = "SELECT * FROM embcuota "
            StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
            StrSql = StrSql & " AND embcuota.embccancela = 0 "
            StrSql = StrSql & " AND embcuota.embcquin = 1 "
            StrSql = StrSql & " AND embcuota.embcanio = " & buliq_periodo!pliqanio
            StrSql = StrSql & " AND embcuota.embcmes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!embcimp
                
                If Cancela Then
                    StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", embccancela = -1 "
                    StrSql = StrSql & ", embcimpreal = " & rs_Cuota!embcimp
                    StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                    StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
                    
            rs_Embargo.MoveNext
        Loop
    Case 3: 'Embargos de la Segunda quincena
        If CBool(USA_DEBUG) Then
            Flog.writeline Espacios(Tabulador * 4) & "Busco los embargos de la segunda quincena "
        End If
    
        StrSql = "SELECT * FROM embargo "
        StrSql = StrSql & " WHERE embargo.ternro = " & buliq_empleado!ternro
        StrSql = StrSql & " AND  embargo.embest = 'E' "
        StrSql = StrSql & " AND embargo.embquincenal = -1 "
        If Nrotpe <> -1 Then
            StrSql = StrSql & " AND embargo.tpenro =" & Nrotpe
        End If
        OpenRecordset StrSql, rs_Embargo
        
        If CBool(USA_DEBUG) Then
            If rs_Embargo.EOF Then
                Flog.writeline Espacios(Tabulador * 4) & "No se encontraron embargos "
            End If
        End If
        
        Do While Not rs_Embargo.EOF
            If CBool(USA_DEBUG) Then
                Flog.writeline Espacios(Tabulador * 4) & "Busco las cuotas para el embargo " & rs_Embargo!prenro
            End If
            StrSql = "SELECT * FROM embcuota "
            StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
            StrSql = StrSql & " AND embcuota.embccancela = 0 "
            StrSql = StrSql & " AND embcuota.embcquin = 2 "
            StrSql = StrSql & " AND embcuota.embcanio = " & buliq_periodo!pliqanio
            StrSql = StrSql & " AND embcuota.embcmes = " & buliq_periodo!pliqmes
            OpenRecordset StrSql, rs_Cuota
            
            If CBool(USA_DEBUG) Then
                If rs_Cuota.EOF Then
                    Flog.writeline Espacios(Tabulador * 4) & "No se encontraron cuotas "
                End If
            End If
            
            Do While Not rs_Cuota.EOF
                Valor = Valor + rs_Cuota!embcimp
                
                If Cancela Then
                    StrSql = "UPDATE embcuota SET pronro =" & buliq_proceso!pronro
                    StrSql = StrSql & ", embccancela = -1 "
                    StrSql = StrSql & ", embcimpreal = " & rs_Cuota!embcimp
                    StrSql = StrSql & " WHERE embcuota.embcnro = " & rs_Cuota!embcnro
                    StrSql = StrSql & " AND embcuota.embnro = " & rs_Cuota!embnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                Bien = True
                
                rs_Cuota.MoveNext
            Loop
            
            rs_Embargo.MoveNext
        Loop
    End Select
            
    ' Chequeo si se liquido la ultima cuota (yestan todas liquidadas).
    ' Si ocurre esto, marco como "Finalizado" el embargo
    StrSql = "SELECT * FROM embcuota "
    StrSql = StrSql & " WHERE embcuota.embnro =" & rs_Embargo!embnro
    StrSql = StrSql & " AND embcuota.embccancela = 0 "
    OpenRecordset StrSql, rs_Aux_Cuota
    If rs_Aux_Cuota.EOF Then
        StrSql = "UPDATE embargo SET embest = 'F'"
        StrSql = StrSql & " WHERE embnro = " & rs_Embargo!embnro
        objConn.Execute StrSql, , adExecuteNoRecords
    End If
End If

If Valor <> 0 Then Valor = -Valor
' cierro todo y libero
'If Param_cur.State = adStateOpen Then Param_cur.Close
If rs_Cuota.State = adStateOpen Then rs_Cuota.Close
If rs_Embargo.State = adStateOpen Then rs_Embargo.Close
        
'Set Param_cur = Nothing
Set rs_Embargo = Nothing
Set rs_Cuota = Nothing

End Sub

