Attribute VB_Name = "mdlConversiones"
Option Explicit

Dim objBTurno As New BuscarTurno
Dim objBDia As New BuscarDia
Dim objFeriado As New Feriado
Dim objFechasHoras As New FechasHoras

Private Sub initVariablesTurno(ByRef T As BuscarTurno)
   p_turcomp = T.Compensa_Turno
   Nro_Grupo = T.Empleado_Grupo
   Nro_Justif = T.Justif_Numero
   justif_turno = T.justif_turno
   Tiene_Justif = T.Tiene_Justif
   Fecha_Inicio = T.FechaInicio
   Nro_fpgo = T.Numero_FPago
   Nro_Turno = T.Turno_Numero
   tiene_turno = T.tiene_turno
   Tipo_Turno = T.Turno_Tipo
   P_Asignacion = T.Tiene_PAsignacion
   
End Sub
Private Sub initVariablesDia(ByRef D As BuscarDia)
   Dia_Libre = D.Dia_Libre
   Nro_Dia = D.Numero_Dia
   Nro_Subturno = D.SubTurno_Numero
   Orden_Dia = D.Orden_Dia
   Trabaja = D.Trabaja
End Sub


Public Sub AD_01(A_HorNro As Long, thdestino As Long, Horas As Single, ByRef Hora_Dest As Single)
'------------------------------------------------------------------------------
'  Descripción: Control de la configuracion de las horas a convertir
'  Autor :
'  Creado:
'  Ult Modif:
'------------------------------------------------------------------------------
Dim Horas_Oblig As Single
Dim objRsCFG As New ADODB.Recordset
Dim objrsAus As New ADODB.Recordset

Hora_Dest = 0
'busco la configuración del tipo de hora a convertir

'FGZ - 24/07/2008 - cambié el query, le saqué las condiciones de minimo y maximo dado que se controlan dentro del loop -------
'StrSql = "SELECT hd_programa,hd_minimo,hd_maximo,hd_excede,hd_fijo,hd_pje " & _
'         " FROM gti_config_horadia " & _
'         " WHERE turnro       = " & Nro_Turno & _
'         " AND   hd_thorigen  = " & A_HorNro & _
'         " AND   hd_thdestino = " & thdestino & _
'         " AND   hd_minimo   <= " & horas & _
'         " AND   hd_maximo   >= " & horas & _
'         " ORDER BY hd_nro"

StrSql = "SELECT hd_programa,hd_minimo,hd_maximo,hd_excede,hd_fijo,hd_pje "
StrSql = StrSql & " FROM gti_config_horadia "
StrSql = StrSql & " WHERE turnro       = " & Nro_Turno
StrSql = StrSql & " AND   hd_thorigen  = " & A_HorNro
StrSql = StrSql & " AND   hd_thdestino = " & thdestino
'FGZ - 15/12/2008 - le agregué estos controles extras
StrSql = StrSql & " AND ( hd_programa is null OR hd_programa = '')"
StrSql = StrSql & " ORDER BY hd_nro"
'FGZ - 24/07/2008 - cambié el query, le saqué las condiciones de minimo y maximo dado que se controlan dentro del loop -------
OpenRecordset StrSql, objRsCFG

If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_01 - busco config de hs a convertir"
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "SQL:" & StrSql
End If

Do While Not objRsCFG.EOF
     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Config para hora : " & thdestino
    ' Si hay un programa de conversion, lo ejecuta
    If objRsCFG!hd_programa <> "" Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tiene programa de conversion asociado"
        End If
        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
    End If

    'Conversion Estandar
    If Horas < objRsCFG!hd_minimo Then
        Hora_Dest = 0
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de horas menor al minimo configurado"
        End If
        Exit Sub
    End If
    
    If Horas > objRsCFG!hd_maximo Then
        If Not objRsCFG!hd_excede Then
            Hora_Dest = 0
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de horas mayor al maximo configurado"
            End If
            Exit Sub
        Else
            If objRsCFG!hd_fijo > 0 Then ' usa fijo
                Hora_Dest = objRsCFG!hd_fijo
                Exit Sub
            Else ' usa Porcentaje
                Hora_Dest = Horas * objRsCFG!hd_pje / 100
                Exit Sub
            End If
        End If
    End If
    
    If objRsCFG!hd_fijo > 0 Then
        Hora_Dest = objRsCFG!hd_fijo
    Else
        Hora_Dest = Horas * objRsCFG!hd_pje / 100
    End If
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Las horas se convertiran a tipo " & Hora_Dest
    End If

    objRsCFG.MoveNext
Loop

FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_01 - Fin"
    End If
'Cierro y libero
    If objRsCFG.State = adStateOpen Then objRsCFG.Close
    If objrsAus.State = adStateOpen Then objrsAus.Close
    Set objRsCFG = Nothing
    Set objrsAus = Nothing
End Sub


Public Sub AD_02(ByVal traza As Boolean, ByVal P_NroTer As Long, ByVal P_NroTur As Long, ByVal P_NroGru As Long, ByVal P_HorNro As Long, ByVal Fecha_Inicio As Date)
'------------------------------------------------------------------------------------------
'Descripcion:   Indica si Convierte Horas o No. Si convierte, indica el Tipo de Hora al cual convierte.
'Autor:         FGZ
'Fecha:         24/08/2006
'Ult. Modif:    reescrita
'------------------------------------------------------------------------------------------
'Dim StrSql As String
Dim thdestino As Long
Dim creado As Boolean
Dim horas_conv As Single

Dim objrsAus As New ADODB.Recordset
Dim objRsConfig As New ADODB.Recordset
Dim objrsHC As New ADODB.Recordset
Dim objrsHoras_C As New ADODB.Recordset
Dim TotHorHHMM As String

Dim acumula As Boolean  'Indica si va acumulando horas o si va reg. a reg.
                        'Esta variable estaba global pero redefinida en el prc01
                        'osea lo que se modificaba aca, cuando salia perdia alcance ==>
                        'No servia para nada
                  
                  
If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_02 - Conversion Antes de Autorizar"
End If

creado = False
horas_conv = 0

'Busca el Horario cumplido
StrSql = "SELECT thnro,horas,horcant FROM gti_horcumplido WHERE (hornro = " & P_HorNro & ")"
OpenRecordset StrSql, objrsHC
If Not objrsHC.EOF Then
    acumula = True
    'Usa_Conv = False
    ''politica que indica si usa la conversion de horas
    'Call Politica(810)
    
    tiene_turno = False
    If Not IsNull(P_NroTur) And (P_NroTur <> 0) Then
        tiene_turno = True
    Else
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No tiene turno"
        End If
        'Exit Sub
        GoTo FIN
    End If
    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = CnTraza
    esFeriado = objFeriado.Feriado(p_fecha, P_NroTer, traza)
    
    If tiene_turno Then
         Set objBDia.Conexion = objConn
         Set objBDia.ConexionTraza = CnTraza
         objBDia.Buscar_Dia p_fecha, Fecha_Inicio, P_NroTur, P_NroTer, P_Asignacion, depurar
         initVariablesDia objBDia
    End If
    
    'Busca si el Tipo de Hora tiene alguna configuracion
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Busca si el Tipo de Hora tiene alguna configuracion"
    End If
    'FGZ - 15/12/2008 - le agregue unos controles extras
    'StrSql = "SELECT hd_thdestino FROM gti_config_horadia WHERE hd_thorigen = " & objrsHC!thnro & " AND  turnro = " & P_NroTur & " ORDER BY hd_nro"
    StrSql = "SELECT hd_thdestino FROM gti_config_horadia WHERE hd_thorigen = " & objrsHC!thnro & " AND  turnro = " & P_NroTur
    StrSql = StrSql & " AND ( hd_programa is null OR hd_programa = '')"
    StrSql = StrSql & " ORDER BY hd_nro"
    OpenRecordset StrSql, objRsConfig
    
    'If Not Usa_Conv Or objRsConfig.EOF Then
    If Not Usa_Conv And Not objRsConfig.EOF Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Realiza Conversion de Horas para el Tipo de Hora Nro: " & objrsHC!thnro
        End If
    
'        StrSql = "SELECT * FROM " & TTempWFAd & " WHERE thnro = " & objrsHC!thnro
'        StrSql = StrSql & " AND Acumula = 0"
'        OpenRecordset StrSql, objrsAus
'        If objrsAus.EOF Then
'            'inserto
'            StrSql = "INSERT INTO " & TTempWFAd & "(thnro,Cant_hs,Acumula) VALUES ("
'            StrSql = StrSql & objrsHC!thnro
'            StrSql = StrSql & "," & objrsHC!horcant & ",0)"
'        Else
'            'Actualizo
'            StrSql = "UPDATE " & TTempWFAd & " SET Cant_hs = " & objrsAus!Cant_hs + objrsHC!horcant
'            'FGZ - 25/04/2007 ---------
'            'StrSql = StrSql & " WHERE thnro = " & objrsHoras_C!hd_thdestino
'            StrSql = StrSql & " WHERE thnro = " & objrsHC!thnro 'objrsHoras_C!hd_thdestino
'            'FGZ - 25/04/2007 ---------
'        End If
        StrSql = "SELECT * FROM " & TTempWFAd & " WHERE thnro = " & objRsConfig!hd_thdestino
        StrSql = StrSql & " AND Acumula = -1"
        OpenRecordset StrSql, objrsAus
        If objrsAus.EOF Then
            'inserto
            StrSql = "INSERT INTO " & TTempWFAd & "(thnro,horas,Cant_hs,Acumula) VALUES ("
            StrSql = StrSql & objRsConfig!hd_thdestino
            StrSql = StrSql & ",'" & objrsHC!Horas & "'"
            StrSql = StrSql & "," & objrsHC!horcant & ",-1)"
        Else
            'Actualizo
            Call SHoras(objrsAus!Horas, objrsHC!Horas, TotHorHHMM)
            StrSql = "UPDATE " & TTempWFAd & " SET horas = '" & TotHorHHMM & "',Cant_hs = " & objrsAus!Cant_hs + objrsHC!horcant
            'FGZ - 25/04/2007 ---------
            'StrSql = StrSql & " WHERE thnro = " & objrsHoras_C!hd_thdestino
            StrSql = StrSql & " WHERE thnro = " & objRsConfig!hd_thdestino
            'FGZ - 25/04/2007 ---------
        End If
        objConn.Execute StrSql, , adExecuteNoRecords
        'Exit Sub
        GoTo FIN
    End If
    
    'Busca la Config. del tipo de hora a convertir
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Busca la Config. del tipo de hora a convertir"
    End If
    StrSql = "SELECT hd_thorigen, hd_thdestino,hd_feriados,hd_nolaborables,hd_laborable,hd_franja FROM gti_config_horadia WHERE turnro = " & P_NroTur
    StrSql = StrSql & " AND hd_thorigen = " & objrsHC!thnro
    StrSql = StrSql & " AND hd_ordenproc = 1 "
    'FGZ - 15/12/2008 - le agregué estos controles
    StrSql = StrSql & " AND ( hd_programa is null OR hd_programa = '')"
    StrSql = StrSql & " ORDER BY hd_nro "
    OpenRecordset StrSql, objrsHoras_C
    Do While Not objrsHoras_C.EOF
        thdestino = objrsHoras_C!hd_thdestino
        If esFeriado And Not objrsHoras_C!hd_feriados Then
            GoTo Siguiente
        Else
            If Dia_Libre And Not objrsHoras_C!hd_nolaborables Then
                GoTo Siguiente
            Else
                If Not Dia_Libre And Not objrsHoras_C!hd_laborable Then
                    GoTo Siguiente
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * (Nivel_Tab_Log + 1)) & "Conversion del Tipo de Hora " & objrsHoras_C!hd_thorigen & " al Tipo de Hora " & objrsHoras_C!hd_thdestino
                    End If
                    acumula = Not objrsHoras_C!hd_franja
                    If acumula Then
                        StrSql = "SELECT * FROM " & TTempWFAd & " WHERE (thnro = " & objrsHoras_C!hd_thdestino & " AND Acumula = -1)"
                        OpenRecordset StrSql, objrsAus
                        
                        If objrsAus.EOF Then
                            'inserto el registro
                            StrSql = "INSERT INTO " & TTempWFAd & "(thnro,horas,Cant_hs,Acumula) VALUES ( "
                            StrSql = StrSql & objrsHoras_C!hd_thdestino & ","
                            StrSql = StrSql & "'" & objrsHC!Horas & "',"
                            StrSql = StrSql & objrsHC!horcant & ","
                            StrSql = StrSql & "-1)"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            
                            creado = True
                        Else
                            If Not creado Then
                                'modifico el campo de cantidad de horas
                                Call SHoras(objrsAus!Horas, objrsHC!Horas, TotHorHHMM)
                                
                                StrSql = "UPDATE " & TTempWFAd & " SET horas = '" & TotHorHHMM & "',Cant_hs = " & objrsAus!Cant_hs + objrsHC!horcant
                                StrSql = StrSql & " WHERE thnro = " & objrsHoras_C!hd_thdestino
                                StrSql = StrSql & " AND Acumula = -1"
                                objConn.Execute StrSql, , adExecuteNoRecords

                                creado = True
                            End If
                        End If
                    Else
                        'Valida y convierte la cantidad
                        Call AD_01(P_HorNro, thdestino, objrsHC!horcant, horas_conv)
                        
                        TotHorHHMM = CHoras(horas_conv, 60)
                        
                        'inserto el registro
                        StrSql = "INSERT INTO " & TTempWFAd & "(thnro,horas,Cant_hs,Acumula) VALUES ( "
                        StrSql = StrSql & objrsHoras_C!hd_thdestino & ","
                        StrSql = StrSql & TotHorHHMM & ","
                        StrSql = StrSql & horas_conv & ","
                        StrSql = StrSql & "-1)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
        
Siguiente:
        creado = False
        objrsHoras_C.MoveNext
    Loop
End If

FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_02 - Fin"
    End If
'Cierro y libero
    If objrsHC.State = adStateOpen Then objrsHC.Close
    If objrsHoras_C.State = adStateOpen Then objrsHoras_C.Close
    If objrsAus.State = adStateOpen Then objrsAus.Close
    If objRsConfig.State = adStateOpen Then objRsConfig.Close
    
    Set objrsHC = Nothing
    Set objrsHoras_C = Nothing
    Set objrsAus = Nothing
    Set objRsConfig = Nothing
End Sub


Public Sub AD_03(ByVal Fecha As Date, ByVal NroTer As Long, ByVal Nro_Grupo As Long, ByVal Nro_Turno As Long)
'------------------------------------------------------------------------------------------
'Descripcion:   Chequea si esta activa la politica 830.
'Autor:         FGZ
'Fecha:
'Ult. Modif:    25/08/2006 - Deshabilitada
'------------------------------------------------------------------------------------------
'FGZ - 12/06/2007 - Esta definiciones hacen macanas dado que estan definidas como globales
'Dim tiene_turno     As Boolean
'Dim Nro_fpgo        As Integer
'Dim Tiene_Justif    As Boolean
'Dim Nro_Justif      As Integer
'Dim justif_turno    As Boolean
'Dim turcomp         As Integer
'Dim Fecha_Inicio    As Date  '/* fecha de inicio del turno */
'Dim P_Asignacion As Boolean
'Dim depurar As Boolean
'
'P_Asignacion = False
'FGZ - 12/06/2007 - Esta definiciones hacen macanas dado que estan definidas como globales

Call Politica(830)
End Sub



Public Sub AD_05(hornro As Long, THOrigen As Long, nrotur As Long, NroTer As Long, Nro_Grupo As Long, Fecha_Inicio As Date, depurar As Boolean)
'------------------------------------------------------------------------------
'Archivo:   gtiad05.p
'Descripci¢n: Convierte los Tipo de Horas que se acumulaban. La conversion de
'             los tipos de horas que se hace reg. a reg. se fue generando sola.
'Input Parameter:
'Output Parameter:
'Autor : Marchese, Juan M.
'Creado: 19/9/2000
'Ult Modif: FGZ - 25/08/2006 - la reescribí toda
'Ult Modif: FGZ - 18/04/2007 - Le saqué este delete .... no se porque estaba pero estaba borrando justamente las horas que generó"
'------------------------------------------------------------------------------
Dim Horas_Acum As Single
Dim linea As String
Dim acumula As Boolean 'Indica si va acumulando horas o si va reg. a reg.
'Dim Usa_Conv As Boolean
Dim horas_conv As Single
Dim creado As Boolean
Dim Horas_Acum_HS As String


'FGZ - 12/06/2007 - Esta definiciones hacen macanas dado que estan definidas como globales
'Dim Trabaja As Boolean
'Dim Orden_Dia As Integer
'Dim Nro_Dia As Integer
'Dim Nro_Subturno As Integer
'Dim Dia_Libre As Boolean
'Dim esFeriado As Boolean
'Dim tiene_turno As Boolean
'Dim Tiene_Justif As Boolean
'Dim Nro_Justif As Boolean
'Dim p_turcomp As Boolean
'Dim Nro_fpgo As Integer
'Dim P_Asignacion As Boolean
'FGZ - 12/06/2007 - Esta definiciones hacen macanas dado que estan definidas como globales

Dim objrsAus As New ADODB.Recordset
Dim objrsHC As New ADODB.Recordset
Dim objrsCFG_H As New ADODB.Recordset

Dim thdestino As Long
Dim TotHorHHMM As String

horas_conv = 0
Horas_Acum = 0
creado = False
Trabaja = False
Dia_Libre = False
esFeriado = False
tiene_turno = False
P_Asignacion = False


If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_05 - Conversion antes de autorizar. Acumulados"
End If

'Busca el Horario cumplido
StrSql = "SELECT hornro FROM gti_horcumplido WHERE (hornro = " & hornro & ")"
OpenRecordset StrSql, objrsHC
If Not objrsHC.EOF Then
    tiene_turno = False
    If Not IsNull(nrotur) And (nrotur <> 0) Then
        tiene_turno = True
    End If
    
    If Not tiene_turno Then
        'Exit Sub
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No tiene turno"
        End If
        GoTo FIN
    End If
    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = CnTraza
    esFeriado = objFeriado.Feriado(p_fecha, NroTer, depurar)
    If tiene_turno Then
         Set objBDia.Conexion = objConn
         Set objBDia.ConexionTraza = CnTraza
         objBDia.Buscar_Dia p_fecha, Fecha_Inicio, nrotur, NroTer, P_Asignacion, depurar
         initVariablesDia objBDia
    End If
    
    'Busca si el Tipo de Hora tiene alguna configuracion
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Busca si el Tipo de Hora tiene alguna configuracion"
    End If
    StrSql = "SELECT hd_thdestino,hd_feriados,hd_nolaborables,hd_laborable FROM gti_config_horadia "
    StrSql = StrSql & " WHERE turnro = " & nrotur
    StrSql = StrSql & " AND hd_thorigen = " & THOrigen
    StrSql = StrSql & " AND hd_franja = 0"
    StrSql = StrSql & " AND hd_ordenproc = 1"
    'FGZ - 15/12/2008 - le agregué estos controles extras
    StrSql = StrSql & " AND ( hd_programa is null OR hd_programa = '')"
    StrSql = StrSql & " ORDER BY hd_nro"
    OpenRecordset StrSql, objrsCFG_H
    Do While Not objrsCFG_H.EOF
    
        If esFeriado And Not objrsCFG_H!hd_feriados Then
            GoTo Siguiente
        Else
            If Dia_Libre And Not objrsCFG_H!hd_nolaborables Then
                GoTo Siguiente
            Else
                If Not Dia_Libre And Not objrsCFG_H!hd_laborable Then
                    GoTo Siguiente
                End If
                
                Horas_Acum = 0
                Horas_Acum_HS = "00:00"
                thdestino = objrsCFG_H!hd_thdestino
                
                'Busco las que todavia no se contemplearon
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Busco las que todavia no se contemplearon"
                End If
                StrSql = "SELECT * FROM " & TTempWFAd
                StrSql = StrSql & " WHERE thnro = " & thdestino
                StrSql = StrSql & " AND Acumula = -1"
                OpenRecordset StrSql, objrsAus
                Do While Not objrsAus.EOF
                    Horas_Acum = Horas_Acum + objrsAus!Cant_hs
                    Call SHoras(Horas_Acum_HS, IIf(IsNull(objrsAus!Horas), "00:00", objrsAus!Horas), Horas_Acum_HS)
                    
                    objrsAus.MoveNext
                Loop
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion tipo de hora " & THOrigen & " a tipo de hora " & thdestino & " Cantidad = " & Horas_Acum
                End If
                
                'Valida y convierte la cantidad
                Call AD_01(THOrigen, thdestino, Horas_Acum, horas_conv)
                
                If horas_conv <> 0 Then
                        TotHorHHMM = CHoras(horas_conv, 60)
                        
                        StrSql = "INSERT INTO " & TTempWFAd & "(thnro,horas,Cant_hs,Acumula) VALUES ("
                        StrSql = StrSql & thdestino & ","
                        StrSql = StrSql & "'" & Horas_Acum_HS & "',"
                        StrSql = StrSql & horas_conv & ","
                        StrSql = StrSql & "0)"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inserto en el temporal TTempWFAd: " & horas_conv
                        End If
                End If
            End If
        End If
        
Siguiente:
    objrsCFG_H.MoveNext
    Loop
End If

'FGZ - 18/04/2007 - Le saqué este delete .... no se porque estaba pero estaba borrando justamente las horas que generó"
'Se borran los wf-ad usados en la generacion
If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Se borran los wf-ad usados en la generacion"
End If
StrSql = "DELETE FROM " & TTempWFAd
StrSql = StrSql & " WHERE acumula = 0"
'objConn.Execute StrSql, , adExecuteNoRecords
'FGZ - 18/04/2007 - Le saqué este delete .... no se porque estaba pero estaba borrando justamente las horas que generó"
      
FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_05 - Fin"
    End If
'Cierro y libero
    If objrsHC.State = adStateOpen Then objrsHC.Close
    If objrsCFG_H.State = adStateOpen Then objrsCFG_H.Close
    If objrsAus.State = adStateOpen Then objrsAus.Close
    
    Set objrsHC = Nothing
    Set objrsCFG_H = Nothing
    Set objrsAus = Nothing
End Sub

Public Sub AD_06(p_fecha As Date, NroTer As Long, depura As Boolean)
'------------------------------------------------------------------------------
'Archivo:   gtiad05.p
'  Descripción: Convierte los Tipo de Horas despues de autorizar
'  Autor :
'  Creado:
'  Ult Modif: FGZ - 06/04/2006
'               no acumulo nada porque sino hace macanas con el sub crearAD
'------------------------------------------------------------------------------
  Dim acumula As Boolean  'Indica si va acumulando horas o si va reg. a reg. */
  'Dim Usa_Conv As Boolean
  Dim horas_conv As Single
  Dim creado As Boolean
  Dim Continua As Boolean
  
  'FGZ - 12/06/2007 - Esta definiciones hacen macanas dado que estan definidas como globales
  'Dim Trabaja As Boolean
  'Dim Orden_Dia As Integer
  'Dim Nro_Dia As Integer
  'Dim Nro_Subturno As Integer
  'Dim Dia_Libre As Boolean
  'Dim esFeriado As Boolean
  'FGZ - 12/06/2007 - Esta definiciones hacen macanas dado que estan definidas como globales
  
  Dim I As Integer
  Dim TotHorHHMM As String
  
  Dim objrsAC_D As New ADODB.Recordset
  Dim objRsConfig As New ADODB.Recordset
  Dim objrsWFAd As New ADODB.Recordset
  
    creado = False
    Trabaja = False
    Dia_Libre = False
    esFeriado = False
    tiene_turno = False
    P_Asignacion = False
    I = 1
    
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion despues de autorizar - AD_06() - Inicio"
    End If
    Call Limpia

    Set objBTurno.Conexion = objConn
    Set objBTurno.ConexionTraza = CnTraza
    objBTurno.Buscar_Turno p_fecha, NroTer, depurar
    initVariablesTurno objBTurno
    If Not tiene_turno And Not Tiene_Justif Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No tiene turno"
        End If
        'Exit Sub
        GoTo FIN
    End If
    
    acumula = True
    'Usa_Conv = False
    ' 'Política que indica si usa la Conversi¢n de Horas
    ' Call Politica(810)
    
    'Chequea si el d¡a para el empleado en feriado o no
    tiene_turno = False
        
    If Not IsNull(Nro_Turno) And (Nro_Turno <> 0) Then
        tiene_turno = True
    End If
    
    If Not tiene_turno Then
        'Exit Sub
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No tiene turno"
        End If
    End If
    
    
    Set objFeriado.Conexion = objConn
    Set objFeriado.ConexionTraza = CnTraza
    esFeriado = objFeriado.Feriado(p_fecha, NroTer, depurar)

   'Busca el dia del empleado, es decir si es un d¡a laborable o no
    If tiene_turno Then
         Set objBDia.Conexion = objConn
         Set objBDia.ConexionTraza = CnTraza
         objBDia.Buscar_Dia p_fecha, Fecha_Inicio, Nro_Turno, NroTer, P_Asignacion, depurar
         initVariablesDia objBDia
    End If
    
    StrSql = "SELECT thnro, adcanthoras FROM gti_acumdiario WHERE ternro = " & NroTer & " AND adfecha = " & ConvFecha(p_fecha)
    OpenRecordset StrSql, objrsAC_D
    Do While Not objrsAC_D.EOF
        'Busca la Config. del tipo de hora a convertir
        StrSql = "SELECT hd_thorigen,hd_thdestino, hd_feriados,hd_nolaborables,hd_laborable,hd_programa "
        StrSql = StrSql & " FROM gti_config_horadia "
        StrSql = StrSql & " WHERE turnro       = " & Nro_Turno
        StrSql = StrSql & " AND   hd_thorigen  = " & objrsAC_D!thnro
        'StrSql = StrSql & " AND   hd_minimo   <= " & objrsAC_D!adcanthoras
        'StrSql = StrSql & " AND   hd_maximo   >= " & objrsAC_D!adcanthoras
        StrSql = StrSql & " AND   hd_ordenproc = 2 "
        'FGZ - 15/12/2008 - le agregué unos controles extras
        StrSql = StrSql & " AND ( hd_programa is null OR hd_programa = '')"
        StrSql = StrSql & " ORDER BY hd_nro"
        ' hd_ordenproc = 2 ==> despues de la autorizaciòn
        ' hd_ordenproc = 1 ==> Antes de la autorizaciòn
        OpenRecordset StrSql, objRsConfig
    
        Do While Not objRsConfig.EOF
'            If esFeriado And Not objRsConfig!hd_feriados Then
'                GoTo SiguienteConfig
'            Else
'                If Dia_Libre And Not objRsConfig!hd_nolaborables Then
'                    GoTo SiguienteConfig
'                Else
'                    If Not Dia_Libre And Not objRsConfig!hd_laborable Then
'                        GoTo SiguienteConfig
'                    End If
'                End If
'            End If
            'FGZ - 10/07/2008 ----------------
            Continua = False
            If esFeriado Then
                If objRsConfig!hd_feriados Then
                    Continua = True
                End If
            Else
                If Dia_Libre And objRsConfig!hd_nolaborables Then
                    Continua = True
                End If
                If Not Dia_Libre And objRsConfig!hd_laborable Then
                    Continua = True
                End If
            End If
            If Not Continua Then
                GoTo SiguienteConfig
            End If
            'FGZ - 10/07/2008 ----------------
            
            If objRsConfig!hd_programa <> "" Then GoTo SiguienteConfig
        
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion tipo de hora " & objRsConfig!hd_thorigen & " a tipo de hora " & objRsConfig!hd_thdestino
            End If
            horas_conv = 0
            Call AD_01(objRsConfig!hd_thorigen, objRsConfig!hd_thdestino, objrsAC_D!adcanthoras, horas_conv)
            If horas_conv <> 0 Then
                
                
                StrSql = "SELECT * FROM " & TTempWFAd
                StrSql = StrSql & " WHERE thnro = " & objRsConfig!hd_thdestino
                StrSql = StrSql & " AND Acumula = -1"
                OpenRecordset StrSql, objrsWFAd
                If objrsWFAd.EOF Then
                    'inserto
                    TotHorHHMM = CHoras(horas_conv, 60)
                    StrSql = "INSERT INTO " & TTempWFAd & "(thnro,horas,Cant_hs,Acumula) VALUES ("
                    StrSql = StrSql & objRsConfig!hd_thdestino & "," & TotHorHHMM & "," & horas_conv & ",-1)"
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log + 1) & " horas a insertar  " & horas_conv & " de tipo de hora " & objRsConfig!hd_thdestino
                    End If
                    
                Else
                    
                    'If Not creado Then
                        ' modifico el campo de cantidad de horas
                        TotHorHHMM = CHorasSF(horas_conv, 60)
                        Call SHoras(TotHorHHMM, IIf(IsNull(objrsWFAd!Horas), "00:00", objrsWFAd!Horas), TotHorHHMM)
                        
                        StrSql = "UPDATE " & TTempWFAd & " SET horas ='" & TotHorHHMM & "',Cant_hs = " & objrsWFAd!Cant_hs + horas_conv
                        StrSql = StrSql & " WHERE thnro = " & objRsConfig!hd_thdestino
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        creado = True
                    'End If
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log + 1) & " horas a agregar  " & horas_conv & " de tipo de hora " & objRsConfig!hd_thdestino
                    End If
                End If
            End If
        
SiguienteConfig:
        objRsConfig.MoveNext
        Loop
    
SiguienteAC_D:
    objrsAC_D.MoveNext
    Loop
    
'If depurar Then
'    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "CrearAD - Inicio"
'End If
Call CrearAD(NroTer)
'If depurar Then
'    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "CrearAD - Fin"
'    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Limpia"
'End If
Call Limpia

FIN:
    If depurar Then
        'Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_06 - Fin"
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion despues de autorizar - AD_06() - FIN"
    End If
'Cierro y libero
    If objrsAC_D.State = adStateOpen Then objrsAC_D.Close
    If objRsConfig.State = adStateOpen Then objRsConfig.Close
    If objrsWFAd.State = adStateOpen Then objrsWFAd.Close
    
    Set objrsAC_D = Nothing
    Set objRsConfig = Nothing
    Set objrsWFAd = Nothing
End Sub


Public Sub AD_07_cambio_en_SABADODOMINGO_MV(p_ternro As Long, p_fecha As Date)
'  --------------------------------------------------------------------------------------------------
'  Archivo:   gtiad07.p
'  Descripción: Convierte los Tipo de Horas despues de autorizar que tengan programas de conversion
'  Autor : ??
'  Fecha : ??
'  Ultima Mod: FGZ - 14/03/2007 ----- Agregado de log
'  ------------------
'  Programas Validos:
'  ------------------
'   1.  Conversion              : Conversion estandar a Jornada produccion
'   2.  ConversionProd          : Customizacion para Moño Azul
'   3.  SACO1HORA               : Customizacion para Temaiken
'   4.  REDONDEO                :
'   5.  SABADOS SCHERING        : Customizacion para Schering
'   6.  NormalesEstrada         : Customizacion para Estrada
'   7.  H100DIVINO              : Customizacion para Divino
'   8.  ConvNormales            : Customizacion para ICI
'   9.  Conv50%                 : Customizacion para ICI
'   10. Conv100%                : Customizacion para ICI
'   11. Conv200%                : Customizacion para ICI
'   12. Feriados                : Customizacion para AGD
'   13. Feriados_Estr           : Customizacion para AGD
'   14. Feriados_Trabajados     : Customizacion para AGD
'   15. Feriados_Trabajados_SD  : Customizacion para AGD
'   16. HorasDestajo            : Customizacion para Frig. Gorina
'   17. Adicalmuerzo            : Customizacion para Frig. Gorina
'   18. Peficiencia             : Customizacion para Frig. Gorina
'   19. Completar               : Customizacion para Schneider.
'   20. SABADODOMINGO MV        : Customizacion para MultiVoice.
'  --------------------------------------------------------------------------------------------------
'  Ult Modif: CAT - 17/05/2006 Conversion DIVINO SA
'  Ult Modif: FGZ - 25/08/2006
'  Ult Modif: FGZ - 22/09/2006 - Customizacion para Moño Azul (ConversionProd)
'  Ult Modif: FGZ - 22/09/2006 - Customizacion para ICI(las encontré en unos fuentes viejos)
'  Ult Modif: FGZ - 14/11/2006 - Customizacion 12 para AGD()
'  Ult Modif: FGZ - 27/11/2006 - Customizaciones 13, 14 y 15 para AGD()
'  Ult Modif: FGZ - 01/02/2007 - Customizaciones 16, 17 y 18 para Gorina()
'  Ult Modif: Diego Rosso - 21/11/2007 - se agrego la Customizacion 19 para Schneider. Completar()
'  Ult Modif: Diego Rosso - 22/01/2008 - se agrego la Customizacion 20 para MultiVoice. SABADODOMINGO MV()
'  --------------------------------------------------------------------------------------------------
Dim Horas_Oblig As Single
Dim HorasRes As Single
Dim TotHor As Single
Dim Hora_Dest As Single
Dim Nro_Dire As Long
Dim Nro_Ccos As Long
Dim Nro_GSeg As Long
Dim RestoDecimal As Single

Dim EntroAntes11 As Boolean
Dim Total50 As Single
Dim Total100 As Single
Dim TotalHoras As Single

Dim Total_Antes13 As Single
Dim Total_Despues13 As Single

Dim TotalNocturnas As Single
Dim Total150 As Single

Dim TipoHora50 As Long
Dim TipoHora100 As Long
Dim TipoHora150 As Long
Dim TipoNocturna As Integer

Dim TipoHoraNoc100 As Long
Dim TipoHoraNoc150 As Long
Dim TipoHoraFer100 As Long
Dim TipoHoraFer150 As Long

Dim QuedanHs As Boolean
Dim SaldoHS As Single
Dim Dias As Integer
Dim Horas As Integer
Dim Minutos As Integer

Dim Limite1 As String
Dim Limite2 As String

Dim CCosto As Long
Dim Sector As Long
Dim Tenro As Long

Dim Tipos_de_Licencias As String
Dim Hay_Licencia As Boolean
Dim Rs_Justif As New ADODB.Recordset
Dim Rs_Lic As New ADODB.Recordset

Dim objRsCFG As New ADODB.Recordset
Dim objRsAD As New ADODB.Recordset
Dim objRsAD100 As New ADODB.Recordset
Dim objrhest As New ADODB.Recordset
Dim rs_HC As New ADODB.Recordset
Dim rs_AD As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Cab As New ADODB.Recordset
Dim SinConvenio As Boolean
Dim ConvenioAnterior As Boolean
Dim rs_TH As New ADODB.Recordset
Dim THOrigen As Long
Dim rs_ST As New ADODB.Recordset
Dim TH_Anormalidad As Long

If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "AD_07 - Conversion despues de autorizar. Programas "
End If
Nivel_Tab_Log = Nivel_Tab_Log + 1
'21/11/2007 - Diego Rosso
'Agrege  gti_config_horadia.hd_maximo y gti_config_horadia.hd_maximo en el select
'*********************************************************
StrSql = "SELECT gti_config_horadia.hd_thdestino, gti_config_horadia.hd_thorigen, gti_config_horadia.hd_programa, gti_acumdiario.adcanthoras, gti_config_horadia.hd_maximo , gti_config_horadia.hd_minimo " & _
         " FROM gti_config_horadia " & _
         " INNER JOIN gti_acumdiario " & _
         " ON gti_acumdiario.thnro = gti_config_horadia.hd_thorigen " & _
         " WHERE hd_programa is not null " & _
         " AND   turnro = " & Nro_Turno & " AND adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro & " ORDER BY hd_nro"
OpenRecordset StrSql, objRsCFG
Do While Not objRsCFG.EOF
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Programa: " & objRsCFG!hd_programa
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo Hora Origen: " & objRsCFG!hd_thorigen
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo Hora Destino: " & objRsCFG!hd_thdestino
        Flog.writeline
    End If
    
    If UCase(objRsCFG!hd_programa) = UCase("Conversion") Then
        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
        StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
        StrSql = StrSql & " ORDER BY diaorden"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!diacanthoras
        End If
        
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        If Not objRsAD.EOF Then
            If (objRsAD!adcanthoras / Horas_Oblig) < 1 Then
                HorasRes = 1
            Else
                HorasRes = objRsAD!adcanthoras / Horas_Oblig
            End If
        
            StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(objRsAD!adcanthoras / Horas_Oblig, 3)
            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(objRsCFG!adcanthoras / Horas_Oblig, 3) & "," & _
                     CInt(False) & "," & CInt(True) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    'FGZ - 08/09/2006 - Customizacion para Moño Azul
    If UCase(objRsCFG!hd_programa) = UCase("ConversionProd") Then
        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
        
        'Saco la cantidad de horas del primer dia del turno
        StrSql = "SELECT * FROM gti_turno WHERE turnro = " & Nro_Turno
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!turcanthsprod
        End If
        If Horas_Oblig > 0 Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            If Not objRsAD.EOF Then
                HorasRes = objRsAD!adcanthoras / Horas_Oblig
            
                StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(objRsAD!adcanthoras / Horas_Oblig, 3)
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(objRsCFG!adcanthoras / Horas_Oblig, 3) & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion abortada, Cantidad de horas produccion del turno es 0."
            End If
        End If
    End If
    
    'Esta es una conversión que se aplica en TMK
    If UCase(objRsCFG!hd_programa) = UCase("SACO1HORA") Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
    
            Hora_Dest = objRsAD!adcanthoras

            StrSql = " SELECT estrcodext FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 35 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                Nro_Dire = CLng(objrhest!estrcodext)
            End If
            
            StrSql = " SELECT estrcodext FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 5 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                Nro_Ccos = CLng(objrhest!estrcodext)
            End If
            
            StrSql = " SELECT his_estructura.estrnro FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 7 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                Nro_GSeg = CLng(objrhest!estrnro)
            End If
            
            If InStr(1, "560,525,542,547,543,530,536", CStr(Nro_GSeg)) > 0 Then
                Hora_Dest = Hora_Dest 'Si pertene a algunos  de los grupos de seguridad de la lista, no convertir
            Else
                If Nro_Dire = 47 Then '/* Gerencia de gastronomia */
                    If Hora_Dest >= 6 Then
                        Hora_Dest = Hora_Dest - 0.5
                    Else
                        Hora_Dest = Hora_Dest
                    End If
                Else
                    If Nro_Ccos <> 286 Then  ' /* Centro atencion a la visita */
                        If Hora_Dest >= 6 Then
                            Hora_Dest = Hora_Dest - 1
                        Else
                            Hora_Dest = Hora_Dest
                        End If
                    End If
                End If
            End If
            
            If Not objRsAD.EOF Then
            
                StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Hora_Dest
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & Hora_Dest & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
    End If 'esta es una conversion que se aplica en ICI

    'FGZ - 23/09/2004
    If UCase(objRsCFG!hd_programa) = "REDONDEO" Then
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        
        If Not objRsAD.EOF Then
            RestoDecimal = objRsAD!adcanthoras - Fix(objRsAD!adcanthoras)
            
            If RestoDecimal <= 0.25 Then
                HorasRes = Fix(objRsAD!adcanthoras)
            Else
                If RestoDecimal >= 0.251 And RestoDecimal <= 0.75 Then
                    HorasRes = Fix(objRsAD!adcanthoras) + 0.5
                Else
                    HorasRes = Fix(objRsAD!adcanthoras) + 1
                End If
            End If
        
            StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(HorasRes, 3)
            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                     CInt(False) & "," & CInt(True) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If

    'FGZ - 28/06/2005
    If UCase(objRsCFG!hd_programa) = "SABADOS SCHERING" Then
        If Weekday(p_fecha) = vbSaturday Then
        
            TipoHora50 = 1
            TipoHora100 = 2
        
            StrSql = " SELECT * FROM gti_horcumplido "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
            StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
            StrSql = StrSql & " Order BY thnro, hornro"
            If rs_HC.State = adStateOpen Then rs_HC.Close
            OpenRecordset StrSql, rs_HC
            
            If Not rs_HC.EOF Then
                rs_HC.MoveFirst
                
                If CInt(Mid(rs_HC!horhoradesde, 1, 2)) < 11 Then
                    EntroAntes11 = True
                Else
                    EntroAntes11 = False
                End If
            Else
                'esto no se deberia dar
                If depurar Then
                    Flog.writeline "No se encontraron horas"
                End If
            End If
            
            If EntroAntes11 Then
                '==> de 00:00 a 13:00 son al 50%
                '  y de 13:00 a 24:00 son al 100%
           
                StrSql = " SELECT * FROM gti_horcumplido "
                StrSql = StrSql & " WHERE ternro = " & p_ternro
                StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                StrSql = StrSql & " ORDER BY thnro, hornro"
                If rs_HC.State = adStateOpen Then rs_HC.Close
                OpenRecordset StrSql, rs_HC
                
                Total50 = 0
                Total100 = 0
                Do While Not rs_HC.EOF
                    If CInt(Mid(rs_HC!horhoradesde, 1, 2)) <= 13 Then
                        hora_desde = rs_HC!horhoradesde
                        If CInt(Mid(rs_HC!horhorahasta, 1, 2)) <= 13 Then
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total50 = Total50 + (Dias * 24) + (Horas + (Minutos / 60))
                        Else
                            'hora_hasta = "1259"
                            hora_hasta = "1300"
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total50 = Total50 + (Dias * 24) + (Horas + (Minutos / 60))
                            
                            hora_desde = "1300"
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                        End If
                    Else
                        hora_desde = rs_HC!horhoradesde
                        hora_hasta = rs_HC!horhorahasta
                        Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                        Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                    End If
                    
                    rs_HC.MoveNext
                Loop
                        
                'Actualizo las hs destino
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If Not objRsAD.EOF Then
                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total50, 3)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    If Total100 <> 0 Then
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD100
                        
                        If Not objRsAD100.EOF Then
                            StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total100, 3)
                            StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora100 & "," & Round(Total100, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            
            Else
                'se pagan las primeras 5 hs al 50 y si quedan se pagan al 100
                'Total50 = IIf(objRsAD!adcanthoras > 5, 5, objRsAD!adcanthoras)
                'Total100 = IIf(objRsAD!adcanthoras - 5 > 0, objRsAD!adcanthoras - 5, 0)
                
                StrSql = " SELECT * FROM gti_horcumplido "
                StrSql = StrSql & " WHERE ternro = " & p_ternro
                StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND thnro = " & objRsCFG!hd_thorigen
                StrSql = StrSql & " ORDER BY hornro"
                If rs_HC.State = adStateOpen Then rs_HC.Close
                OpenRecordset StrSql, rs_HC
                
                TotalHoras = 0
                Do While Not rs_HC.EOF
                    TotalHoras = TotalHoras + rs_HC!horcant
                
                    rs_HC.MoveNext
                Loop
                Total50 = IIf(TotalHoras > 5, 5, TotalHoras)
                Total100 = IIf(TotalHoras - 5 > 0, TotalHoras - 5, 0)
                If Total100 <> 0 Then
                    QuedanHs = True
                Else
                    QuedanHs = False
                End If
                
                'Actualizo las hs destino
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If Not objRsAD.EOF Then
                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total50, 3)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    If Total100 <> 0 Then
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD100
                        
                        If Not objRsAD100.EOF Then
                            StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total100, 3)
                            StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora100 & "," & Round(Total100, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
    End If
    
    'FGZ - 27/07/2005
    'Nueva Politica de Convrersion para Estrada
    If objRsCFG!hd_programa = "NormalesEstrada" Then
        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        StrSql = StrSql & " ORDER BY diaorden"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!diacanthoras
        End If
        
        StrSql = " SELECT * FROM gti_acumdiario"
        StrSql = StrSql & " INNER JOIN gti_config_tur_hor ON gti_config_tur_hor.thnro = gti_acumdiario.thnro "
        StrSql = StrSql & " WHERE gti_config_tur_hor.turnro = " & Nro_Turno
        StrSql = StrSql & " AND gti_acumdiario.adfecha = " & ConvFecha(p_fecha)
        StrSql = StrSql & " AND gti_acumdiario.ternro = " & p_ternro
        StrSql = StrSql & " AND gti_config_tur_hor.conhornro IN (2,4,5,19)"
        If rs_AD.State = adStateOpen Then rs_AD.Close
        OpenRecordset StrSql, rs_AD
        
        If rs_AD.EOF Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            
            If Not objRsAD.EOF Then
                StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Horas_Oblig
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Horas_Oblig & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    'Nueva Politica de Convrersion para Estrada

    'CAT - 14/02/2006
    If UCase(objRsCFG!hd_programa) = "H100DIVINO" Then
        
'            TipoHora100 = 2
'            TipoHora150 = 44
'
'            TipoHoraNoc100 = 45
'            TipoHoraNoc150 = 46
'            TipoHoraFer100 = 47
'            TipoHoraFer150 = 48

            TipoHora100 = 6
            TipoHora150 = 38
            TipoNocturna = 8

            TipoHoraNoc100 = 40
            TipoHoraNoc150 = 41
            TipoHoraFer100 = 42
            TipoHoraFer150 = 43
            
            StrSql = " SELECT * FROM gti_acumdiario "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
            StrSql = StrSql & " AND thnro = " & TipoHora100
            If rs_AD.State = adStateOpen Then rs_AD.Close
            OpenRecordset StrSql, rs_AD
            If rs_AD.EOF Then
                'No hay horas al 100 -> Nada que hacer
                If depurar Then
                    Flog.writeline "No se encontraron Acumulado de Horas. SQL " & StrSql
                End If
            Else
                Total100 = rs_AD!adcanthoras
                If depurar Then
                    Flog.writeline "Total horas 100: " & Total100
                End If

                'Calculo la cantidad de horas Obligatorias
                StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    Horas_Oblig = objRs!diacanthoras
                End If
    
                If Horas_Oblig < 8 Then
                    'si es frando Horas_Oblig = 0 -> entonces vale 8
                    'si es menor que 8, tiene que ser la diferencia de horas
                    'para llegar a 8
                    Horas_Oblig = 8 - Horas_Oblig
                End If
                
                If depurar Then
                    Flog.writeline "Total Obligatorias: " & Horas_Oblig
                End If
                
                If Total100 > Horas_Oblig Then
                    'si las extras 100 son mas que las obligatorias
                    'convertir el exedente en 150
                    
                    'Actualizo las horas al 100 como la cantidad de horas Obligatorias
                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Horas_Oblig, 3)
                    StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    HorasRes = Total100 - Horas_Oblig
                    
                    'Inserto las horas 150 como el exedente de las horas Obligatorias
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora150 & "," & Round(HorasRes, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                
                'Feriado
                If esFeriado Then
                    'reemplazo las Hs100 por Feriado100 y
                    '              Hs150 por Feriado150
                
                    StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraFer100
                    StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                    StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraFer150
                    StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                
                  'Nocturnas
                  Total150 = HorasRes
                  'TotalNocturnas = ContarHs(P_ternro, p_fecha, "2200", p_fecha + 1, "0600")
                  
                  TotalNocturnas = 0
                  StrSql = " SELECT * FROM gti_acumdiario "
                  StrSql = StrSql & " WHERE ternro = " & p_ternro
                  StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
                  StrSql = StrSql & " AND thnro = " & TipoNocturna
                  If rs_AD.State = adStateOpen Then rs_AD.Close
                  OpenRecordset StrSql, rs_AD
                
                  TotalNocturnas = rs_AD!adcanthoras
                  
                  Do While (HorasRes > 0 Or Total100 > 0) And (TotalNocturnas > 0)
                        If depurar Then
                            Flog.writeline "Total de Horas Nocturnas: " & TotalNocturnas
                            Flog.writeline "Total de Horas 100: " & Total100
                            Flog.writeline "Total de Horas 150: " & Total150
                        End If
                      If Total150 > 0 Then
                          'cambio las 150 por nocturnas 150
                          If TotalNocturnas < Total150 Then
                              StrSql = " UPDATE gti_acumdiario SET adcanthoras = adcanthoras - " & Round(TotalNocturnas, 3)
                              StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Actualizo Horas 150. " & StrSql
                              End If
                          
                              StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                       " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHoraNoc150 & "," & Round(TotalNocturnas, 3) & "," & _
                                       CInt(False) & "," & CInt(True) & ")"
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Inserto Horas Noct 150. " & StrSql
                              End If
                              Total150 = Total150 - TotalNocturnas
                              TotalNocturnas = 0
                          Else
                              If TotalNocturnas = Total150 Then
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc150
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 150 a Noct 150. " & StrSql
                                  End If
                                  Total150 = 0
                                  TotalNocturnas = 0
                              Else    'TotalNocturnas > Total150
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc150
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 150 a Noct 150. " & StrSql
                                  End If
                                  
                                  TotalNocturnas = TotalNocturnas - Total150
                                  Total150 = 0
                              End If
                          End If
                      Else
                          'cambio las 100 por nocturnas 100
                          If TotalNocturnas < Total100 Then
                              StrSql = " UPDATE gti_acumdiario SET adcanthoras = adcanthoras - " & Round(TotalNocturnas, 3)
                              StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Actualizo Horas 100. " & StrSql
                              End If
                              StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                       " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHoraNoc100 & "," & Round(TotalNocturnas, 3) & "," & _
                                       CInt(False) & "," & CInt(True) & ")"
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Insert Horas Noct 100. " & StrSql
                              End If
                          
                              Total100 = Total100 - TotalNocturnas
                              TotalNocturnas = 0
                          Else
                              If TotalNocturnas = Total100 Then
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc100
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 100 a Noct 100. " & StrSql
                                  End If
                                  Total100 = 0
                                  TotalNocturnas = 0
                              Else    'TotalNocturnas > Total100
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc100
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 100 a Noct 100. " & StrSql
                                  End If
                                  TotalNocturnas = TotalNocturnas - Total100
                                  Total100 = 0
                              End If
                          End If
                      End If
                  Loop
               End If
            End If
    End If

    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "ConvNormales" Then
        If Weekday(p_fecha) = 1 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If (rs_HC!horhoradesde > rs_HC!horhorahasta) Then
                    objFechasHoras.RestaHs p_fecha, "0000", p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                    HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                    StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                    StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            
            End If
        End If
        If Weekday(p_fecha) = 7 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If rs_HC!horhoradesde >= Limite1 Then
                
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    
                    If rs_HC!horhorahasta > Limite1 Then
                    
                        objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, Limite1, Tdias, Thoras, Tmin
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                    Else
                        objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If
    
    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "Conv50%" Then
        If Weekday(p_fecha) <> 7 Then
             StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
             StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If Weekday(p_fecha) = 7 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If rs_HC!horhoradesde >= Limite2 Then
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If rs_HC!horhoradesde > Limite1 Then
                        If rs_HC!horhorahasta >= Limite2 Then
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, Limite2, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Else
                        If rs_HC!horhorahasta >= Limite2 Then
                            objFechasHoras.RestaHs p_fecha, Limite1, p_fecha, Limite2, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            objFechasHoras.RestaHs p_fecha, Limite1, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "Conv200%" Then
        If Weekday(p_fecha) <> 1 Then
             StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
             StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
    
    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "Conv100%" Then
        If Weekday(p_fecha) <> 7 And Weekday(p_fecha) <> 1 Then
             StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
             StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If (Weekday(p_fecha) = 1) Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If (rs_HC!horhoradesde > rs_HC!horhorahasta) Then
                    objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, "2400", Tdias, Thoras, Tmin
                    HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                    StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                    StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
        
        If Weekday(p_fecha) = 7 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If rs_HC!horhorahasta <= Limite2 And (rs_HC!horhoradesde < rs_HC!horhorahasta) Then
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If rs_HC!horhoradesde >= Limite2 Then
                        If (rs_HC!horhoradesde < rs_HC!horhorahasta) Then
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        Else
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha + 1, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        End If
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        objFechasHoras.RestaHs p_fecha, Limite2, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        StrSql = "UPDATE gti_acumdiario SET adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If

    'FGZ - 27/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados") Then
        Set objFeriado.Conexion = objConn
        Set objFeriado.ConexionTraza = CnTraza
        esFeriado = objFeriado.Feriado(p_fecha, Empleado.Ternro, depurar)
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Feriado? " & esFeriado
        End If
        
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        If Not objRsAD.EOF Then
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            If esFeriado Then
                'Reviso que no tenga licencia de:
                '   Accidente(9,13,14)
                '   Enfermedad(8)
                '   Casamiento(4)
                '   Vacaciones(2)
                '   Nacimiento(3)
                '   Fallecimiento(5)
                Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                
                'Busco si el dia tiene justificacion
                StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                OpenRecordset StrSql, Rs_Justif
                If Not Rs_Justif.EOF Then
                    'Busco la licencia
                    StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                    StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                    StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                    StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                    StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                    OpenRecordset StrSql, Rs_Lic
                    If Rs_Lic.EOF Then
                        Hay_Licencia = False
                    Else
                        Hay_Licencia = True
                    End If
                Else
                    Hay_Licencia = False
                End If
        
                If Not Hay_Licencia Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Hay Licencia. Se paga."
                    End If
                
                    'Si Origen = destino ==> Quedan como estan
                    '                    Sino Creo el destino
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'Inserto el tipo de hora
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Horas_Oblig & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'Quedan como estan
                    End If
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hay Licencia. No se paga."
                    End If
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'No se pagan
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = 0"
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'No se pagan ==> borro
                        StrSql = " DELETE gti_acumdiario "
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        Else    'El tipo de hora destino no existe
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El tipo de hora destino no existe"
            End If
            
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            If esFeriado Then
                'Reviso que no tenga licencia de:
                '   Accidente(9,13,14)
                '   Enfermedad(8)
                '   Casamiento(4)
                '   Vacaciones(2)
                '   Nacimiento(3)
                '   Fallecimiento(5)
                Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                
                'Busco si el dia tiene justificacion
                StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                OpenRecordset StrSql, Rs_Justif
                If Not Rs_Justif.EOF Then
                    'Busco la licencia
                    StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                    StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                    StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                    StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                    StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                    OpenRecordset StrSql, Rs_Lic
                    If Rs_Lic.EOF Then
                        Hay_Licencia = False
                    Else
                        Hay_Licencia = True
                    End If
                Else
                    Hay_Licencia = False
                End If
        
                If Not Hay_Licencia Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Hay Licencia. Se paga."
                    End If
                    
                    'Si Origen = destino ==> Quedan como estan
                    '                    Sino Creo el destino
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'Inserto el tipo de hora
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Horas_Oblig & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'Quedan como estan
                    End If
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hay Licencia. No se paga."
                    End If
                    
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'no las creo porque no se pagan
                    Else
                        'No se pagan ==> las borro
                        StrSql = " DELETE gti_acumdiario"
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para AGD


    'FGZ - 14/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_Estr") Then
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        If Not objRsAD.EOF Then
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            Set objFeriado.Conexion = objConn
            Set objFeriado.ConexionTraza = CnTraza
            esFeriado = objFeriado.Feriado(p_fecha, Empleado.Ternro, depurar)
            
            If esFeriado Then
                If Not Feriado_Por_Estructura Then
                    'Reviso que no tenga licencia de:
                    '   Accidente(9,13,14)
                    '   Enfermedad(8)
                    '   Casamiento(4)
                    '   Vacaciones(2)
                    '   Nacimiento(3)
                    '   Fallecimiento(5)
                    Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                    
                    'Busco si el dia tiene justificacion
                    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                    OpenRecordset StrSql, Rs_Justif
                    If Not Rs_Justif.EOF Then
                        'Busco la licencia
                        StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                        StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                        StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                        StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                        OpenRecordset StrSql, Rs_Lic
                        If Rs_Lic.EOF Then
                            Hay_Licencia = False
                        Else
                            Hay_Licencia = True
                        End If
                    Else
                        Hay_Licencia = False
                    End If
            
                    If Not Hay_Licencia Then
                        'Si Origen = destino ==> Quedan como estan
                        '                    Sino Creo el destino
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'Inserto el tipo de hora
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Horas_Oblig & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'Quedan como estan
                        End If
                    Else
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'No se pagan
                            StrSql = " UPDATE gti_acumdiario SET adcanthoras = 0"
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'No se pagan ==> borro
                            StrSql = " DELETE gti_acumdiario "
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else    'Feriado por estructura
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'No se pagan
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = 0"
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'No se pagan ==> borro
                        StrSql = " DELETE gti_acumdiario "
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        Else    'El tipo de hora destino no existe
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            If esFeriado Then
                If Not Feriado_Por_Estructura Then
                    'Reviso que no tenga licencia de:
                    '   Accidente(9,13,14)
                    '   Enfermedad(8)
                    '   Casamiento(4)
                    '   Vacaciones(2)
                    '   Nacimiento(3)
                    '   Fallecimiento(5)
                    Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                    
                    'Busco si el dia tiene justificacion
                    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                    OpenRecordset StrSql, Rs_Justif
                    If Not Rs_Justif.EOF Then
                        'Busco la licencia
                        StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                        StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                        StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                        StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                        OpenRecordset StrSql, Rs_Lic
                        If Rs_Lic.EOF Then
                            Hay_Licencia = False
                        Else
                            Hay_Licencia = True
                        End If
                    Else
                        Hay_Licencia = False
                    End If
            
                    If Not Hay_Licencia Then
                        'Si Origen = destino ==> Quedan como estan
                        '                    Sino Creo el destino
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'Inserto el tipo de hora
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Horas_Oblig & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'Quedan como estan
                        End If
                    Else
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'no las creo porque no se pagan
                        Else
                            'No se pagan ==> las borro
                            StrSql = " DELETE gti_acumdiario"
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else    'Feriado por estructura
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'No se pagan
                    Else
                        'No se pagan
                    End If
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para AGD


    'FGZ - 27/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_Trabajados") Then
        'las obligatorias serian las de horigen
        'Las excedentes serian las destino
    
        'Calculo la cantidad de horas Obligatorias
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!diacanthoras
        End If
        If Horas_Oblig < 8 Then
            'si es frando Horas_Oblig = 0 -> entonces vale 8
            'si es menor que 8, tiene que ser la diferencia de horas
            'para llegar a 8
            Horas_Oblig = 8 - Horas_Oblig
        End If
        If depurar Then
            Flog.writeline "Total Obligatorias: " & Horas_Oblig
        End If
        TotHor = objRsCFG!adcanthoras
        If depurar Then
            Flog.writeline "Cant Hs Origen: " & TotHor
        End If
        If TotHor > Horas_Oblig Then
            'Horas Origen
            StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Horas_Oblig
            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Horas excedentes
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            If objRsAD.EOF Then
                'Inserto
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & (TotHor - Horas_Oblig) & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else    'El tipo de hora destino no existe
                'Actualizo el valor
                StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & (TotHor - Horas_Oblig)
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            'Quedan como estan
            
        End If
    End If
    'Nueva Politica de Convrersion para AGD

    'FGZ - 27/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_Trabajados_SD") Then
        'las obligatorias serian las de horigen
        'Las excedentes serian las destino
        
        If Weekday(p_fecha) > 1 And Weekday(p_fecha) < 7 Then 'Dia de semana
            'Calculo la cantidad de horas Obligatorias
            StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Horas_Oblig = objRs!diacanthoras
            End If
            If Horas_Oblig < 8 Then
                'si es frando Horas_Oblig = 0 -> entonces vale 8
                'si es menor que 8, tiene que ser la diferencia de horas
                'para llegar a 8
                Horas_Oblig = 8 - Horas_Oblig
            End If
            If depurar Then
                Flog.writeline "Total Obligatorias: " & Horas_Oblig
            End If
            TotHor = objRsCFG!adcanthoras
            If depurar Then
                Flog.writeline "Cant Hs Origen: " & TotHor
            End If
            If TotHor > Horas_Oblig Then
                'Horas Origen
                StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Horas_Oblig
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
                
                'Horas excedentes
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If objRsAD.EOF Then
                    'Inserto
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & (TotHor - Horas_Oblig) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else    'El tipo de hora destino no existe
                    'Actualizo el valor
                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & (TotHor - Horas_Oblig)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                'Quedan como estan
            End If
        Else
            If Weekday(p_fecha) = 1 Then 'Domingo
                'Se convierten todas las horas al destino
                TotHor = objRsCFG!adcanthoras
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If objRsAD.EOF Then
                    'Inserto
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHor & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else    'El tipo de hora destino no existe
                    'Actualizo el valor
                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & TotHor
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else    'Sabado
            
                StrSql = " SELECT * FROM gti_horcumplido "
                StrSql = StrSql & " WHERE ternro = " & p_ternro
                StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND thnro = " & objRsCFG!hd_thorigen
                StrSql = StrSql & " ORDER BY hornro"
                If rs_HC.State = adStateOpen Then rs_HC.Close
                OpenRecordset StrSql, rs_HC
                
                Total_Antes13 = 0
                Total_Despues13 = 0
                Do While Not rs_HC.EOF
                    If CInt(Mid(rs_HC!horhoradesde, 1, 2)) <= 13 Then
                        hora_desde = rs_HC!horhoradesde
                        If CInt(Mid(rs_HC!horhorahasta, 1, 2)) <= 13 Then
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total_Antes13 = Total_Antes13 + (Dias * 24) + (Horas + (Minutos / 60))
                        Else
                            hora_hasta = "1300"
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total_Antes13 = Total_Antes13 + (Dias * 24) + (Horas + (Minutos / 60))
                            
                            hora_desde = "1300"
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total_Despues13 = Total_Despues13 + (Dias * 24) + (Horas + (Minutos / 60))
                        End If
                    Else
                        hora_desde = rs_HC!horhoradesde
                        hora_hasta = rs_HC!horhorahasta
                        Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                        Total_Despues13 = Total_Despues13 + (Dias * 24) + (Horas + (Minutos / 60))
                    End If
                    
                    rs_HC.MoveNext
                Loop
                
                'Actualizo las hs Origen
                If Total_Antes13 <> 0 Then
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total_Antes13, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thorigen & "," & Round(Total_Antes13, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        'llevo la cantidad a 0
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(0, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'No hago nada
                    End If
                
                End If
                
                If Total_Despues13 <> 0 Then
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total_Despues13, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(Total_Despues13, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para AGD


    'FGZ - 01/02/2007
    'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
    If UCase(objRsCFG!hd_programa) = UCase("HORASDESTAJO") Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HORASDESTAJO -->"
            End If
            'Busco Centro de Costo y el Sector
            CCosto = 0
            Sector = 0
            
            Tenro = 5   'Centro De Costo
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                CCosto = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Centro de Costo."
                End If
            End If
        
            Tenro = 2   'Sector
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Sector = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Sector"
                End If
            End If
        
            If CCosto <> 0 And Sector <> 0 Then
                StrSql = " SELECT ordtcant FROM orden_trabajo "
                StrSql = StrSql & " WHERE estrnro2 = " & Sector
                StrSql = StrSql & " AND (estrnro3 = 0 OR estrnro3 is null OR estrnro3 = " & CCosto & ")"
                StrSql = StrSql & " AND ordtfecdesde = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, rs_Cab
                If Not rs_Cab.EOF Then
                    HorasRes = rs_Cab!ordtcant
                    
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(HorasRes, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No hay ordenes para la fecha: " & StrSql
                    End If
                End If
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "CCosto = 0 OR Sector = 0. " & CCosto & " y " & Sector
                End If
            End If
        End If
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        
    'FGZ - 01/02/2007
    'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
    If UCase(objRsCFG!hd_programa) = UCase("ADICALMUERZO") Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "ADICALMUERZO -->"
            End If
            'Busco Centro de Costo y el Sector
            CCosto = 0
            Sector = 0
            
            Tenro = 5   'Centro De Costo
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                CCosto = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Centro de Costo."
                End If
            End If
        
            Tenro = 2   'Sector
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Sector = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Sector"
                End If
            End If
        
            If CCosto <> 0 And Sector <> 0 Then
                StrSql = " SELECT ordtcant FROM orden_trabajo "
                StrSql = StrSql & " WHERE estrnro2 = " & Sector
                StrSql = StrSql & " AND (estrnro3 = 0 OR estrnro3 is null OR estrnro3 = " & CCosto & ")"
                StrSql = StrSql & " AND ordtfecdesde = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, rs_Cab
                If Not rs_Cab.EOF Then
                    HorasRes = rs_Cab!ordtcant
                    
                    If HorasRes > 400 Then
                        HorasRes = 40
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD
                        If Not objRsAD.EOF Then
                            StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(HorasRes, 3)
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
            End If
        End If
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        
        'FGZ - 01/02/2007
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        If UCase(objRsCFG!hd_programa) = UCase("PEFICIENCIA") Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "PEFICIENCIA -->"
            End If
            'Busco Centro de Costo y el Sector
            CCosto = 0
            Sector = 0
            
            Tenro = 5   'Centro De Costo
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                CCosto = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Centro de Costo."
                End If
            End If
        
            Tenro = 2   'Sector
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Sector = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Sector"
                End If
            End If
        
            If CCosto <> 0 And Sector <> 0 Then
                StrSql = " SELECT ordtcant FROM orden_trabajo "
                StrSql = StrSql & " WHERE estrnro2 = " & Sector
                StrSql = StrSql & " AND (estrnro3 = 0 OR estrnro3 is null OR estrnro3 = " & CCosto & ")"
                StrSql = StrSql & " AND ordtfecdesde = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, rs_Cab
                If Not rs_Cab.EOF Then
                    HorasRes = rs_Cab!ordtcant
                    
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(HorasRes, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        
    'Diego Rosso - 21/11/2007
    'CAS-04990- Nueva Conversión GTI - Schneider Electric -----------------------
    'Completar la cantidad de Hs. faltantes para llegar a totalizar las horas mínimas de jornada
    'en los dias Configurados.
    'Tiene en cuenta el mínimo de Horas Normales realizadas para que esto tenga efecto.
    If UCase(objRsCFG!hd_programa) = UCase("Completar") Then
        
        'Llamo a la politica
        Call Politica(571)
        TH_Anormalidad = st_TipoHora1
        
        If depurar Then
           Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo de Hora de Anormalidad " & TH_Anormalidad
           Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Lista de Dias " & st_ListaTH
           Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Dia de la semama a procesar " & Weekday(p_fecha)
           StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
           OpenRecordset StrSql, objRsAD
           If Not objRsAD.EOF Then
                If depurar Then
                   Flog.writeline Espacios(Tabulador * 4) & "Antes de la conversion hay " & objRsAD!adcanthoras & " horas destino"
                End If
           Else
                If depurar Then
                   Flog.writeline Espacios(Tabulador * 4) & "Antes de la conversion NO hay horas destino"
                End If
           End If
        End If
        
        'Chequear que este configurada la politica
        If Not EsNulo(st_ListaTH) Then
        'If (Not IsNull(st_ListaTH)) Or (st_ListaTH <> "") Then
            'weekday. Chequeo que el dia que estoy procesando este en la lista sino no lo proceso
            If InStr(1, st_ListaTH, Weekday(p_fecha)) > 0 Then
        
                     'objRsCFG!diacanthoras = Cantidad de horas que hizo el empleado
                    'Chequeo que las horas que hizo el empleado esten dentro del minimo y maximo configurado para la conversion
                    If (objRsCFG!adcanthoras <= objRsCFG!hd_maximo) And (objRsCFG!adcanthoras >= objRsCFG!hd_minimo) Then
        
                         'Busca Cantidad de Horas de Día para el turno del Empleado
                         StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
                         OpenRecordset StrSql, objRs
                         If Not objRs.EOF Then
                             Horas_Oblig = objRs!diacanthoras ' Cantidad de horas configuradas para el dia
                         Else
                            Horas_Oblig = 0
                            If depurar Then
                               Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontró la Cantidad de Horas del día " & Nro_Dia
                            End If
                         End If
                         
                         'Cantidad de horas a convertir
                         HorasRes = Horas_Oblig - objRsCFG!adcanthoras
                         If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas Obligatorias: " & Horas_Oblig
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas de origen encontradas: " & objRsCFG!adcanthoras
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas a compensar: " & HorasRes
                         End If
                         
                         'chequea negativo
                         If HorasRes > 0 Then
                            'Revisar si existe ST por esa cantidad de hs
                            '   si existe ==> borro anormalidad y genero compensacion
                            '   sino ... nada
                            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TH_Anormalidad & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            StrSql = StrSql & " AND adcanthoras = " & HorasRes
                            OpenRecordset StrSql, rs_ST
                            If Not rs_ST.EOF Then
                                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                OpenRecordset StrSql, objRsAD
                                If Not objRsAD.EOF Then
                                    If depurar Then
                                       Flog.writeline Espacios(Tabulador * 4) & "Actualizo horas destino --> " & Round(objRsAD!adcanthoras + HorasRes, 3)
                                    End If
                                    'StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(objRsAD!adcanthoras + HorasRes, 3)
                                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(HorasRes, 3)
                                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    If depurar Then
                                       Flog.writeline Espacios(Tabulador * 4) & "Inserto horas destino --> " & Round(HorasRes, 3)
                                    End If
                                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & Round(HorasRes, 3) & "," & _
                                             CInt(False) & "," & CInt(True) & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                                
                                'Borro el tipo de hora del AD
                                StrSql = " DELETE gti_acumdiario WHERE thnro = " & TH_Anormalidad & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                StrSql = StrSql & " AND adcanthoras = " & HorasRes
                                objConn.Execute StrSql, , adExecuteNoRecords
                                
                                'Cambio el la condicion de anormalidad del HC anormalidad
                                StrSql = "UPDATE gti_horcumplido set normnro = 10, normnro2 = 10 WHERE thnro = " & TH_Anormalidad & " AND ternro = " & p_ternro & " AND horfecrep = " & ConvFecha(p_fecha)
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                If depurar Then
                                   Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No compensa porque no encontró la anormalidad. "
                                End If
                            End If
                         Else 'HorasRes > 0
                            If depurar Then
                               Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La cantidad de horas de que registradas en el dia iguala o supera la cantidad de horas configuradas. "
                            End If
                         End If 'HorasRes > 0
                        
                    Else '(objRsCFG!diacanthoras <= objRsCFG!hd_maximo) And (objRsCFG!diacanthoras >= objRsCFG!hd_minimo)
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La cantidad de horas de origen no entran entre el minimo y maximo configurado."
                        End If
                    End If '(objRsCFG!diacanthoras <= objRsCFG!hd_maximo) And (objRsCFG!diacanthoras >= objRsCFG!hd_minimo)
            Else 'InStr(1, st_ListaTH, Weekday(p_fecha)) > 0
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El dia no esta dentro de la lista de dias configurada."
                End If
            End If 'InStr(1, st_ListaTH, Weekday(p_fecha)) > 0
        Else '(Not st_ListaTH Is Null) Or (st_ListaTH <> "")
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La politica 571 no esta configurada."
            End If
        End If '(Not st_ListaTH Is Null) Or (st_ListaTH <> "")
    End If
        
        
    '*********************************************************************************************************************************************************************
    'Diego Rosso - 22/01/2008
    'Sabado hasta las 13 hs  genere Hs adicionales Aut. y después del sabado a las 13 y hasta el domingo genera Hs 100%  según corresponda.
    '*********************************************************************************************************************************************************************
    If UCase(objRsCFG!hd_programa) = "SABADODOMINGO MV" Then
        Tenro = 19   'Convenio
        'Tenro = 55   'Convenio
        SinConvenio = True
        StrSql = " SELECT estructura.estrnro, estructura.estrcodext FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
        StrSql = StrSql & " WHERE his_estructura.ternro = " & p_ternro
        StrSql = StrSql & " AND estructura.tenro =" & Tenro
        StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(p_fecha) & ")"
        StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            If Not EsNulo(rs_Estructura!estrcodext) Then
                ConvenioAnterior = IIf(rs_Estructura!estrcodext = "0", True, False)
                SinConvenio = False
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El Convenio " & rs_Estructura!estrnro & " no tiene configurado el codigo externo. No se ejecutará la coversión."
                End If
            End If
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Convenio. No se ejecutará la coversión."
            End If
        End If
        
        If Not SinConvenio Then
            If ConvenioAnterior Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio Anterior..."
                End If
                TipoHora50 = 3
                TipoHora100 = 7
                If Weekday(p_fecha) = vbSaturday Or Weekday(p_fecha) = vbSunday Then
                    THOrigen = objRsCFG!hd_thorigen
                    
                    'Busco el horario trabajado en el dia
                    StrSql = " SELECT * FROM gti_horcumplido "
                    StrSql = StrSql & " WHERE ternro = " & p_ternro
                    StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                    StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                    StrSql = StrSql & " AND thnro = " & THOrigen
                    StrSql = StrSql & " Order BY thnro, hornro"
                    OpenRecordset StrSql, rs_HC
                    If rs_HC.EOF Then
                        'quiere decir que este tipo de hora fué autorizado ==>
                        'debo buscar el tipo de hora no autorizado y
                        'topearlo a la cantidad de hs autorizadas
                        StrSql = "SELECT thnro FROM tiphora WHERE thautpor = " & THOrigen
                        OpenRecordset StrSql, rs_TH
                        If Not rs_TH.EOF Then
                            THOrigen = rs_TH!thnro

                            StrSql = " SELECT * FROM gti_horcumplido "
                            StrSql = StrSql & " WHERE ternro = " & p_ternro
                            StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                            StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                            StrSql = StrSql & " AND thnro = " & THOrigen
                            StrSql = StrSql & " Order BY thnro, hornro"
                            OpenRecordset StrSql, rs_HC
                        End If
                    End If
                    
                      If depurar Then
                            Flog.writeline "Fecha a procesar:" & p_fecha & " Correspondiente a un dia Sabado"
                      End If
                        'Si es sabado
                        '==> de 00:00 a 13:00 quedan igual
                        '  y de 13:00 a 24:00 son al 100%
                        'Si es Domingo
                        'Todas las horas se pasan al 100%
                        
                        Total100 = 0
                        Total50 = 0
                        Do While Not rs_HC.EOF
                            If CInt(Mid(rs_HC!horhoradesde, 1, 4)) <= 1300 And Weekday(p_fecha) = vbSaturday Then
        '                        hora_desde = rs_HC!horhoradesde
                                If CInt(Mid(rs_HC!horhorahasta, 1, 2)) >= 13 Then
                                    hora_desde = rs_HC!horhoradesde
                                    hora_hasta = "1300"
                                    Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                                    Total50 = Total50 + (Dias * 24) + (Horas + (Minutos / 60))
                                    
                                    hora_desde = "1300"
                                    hora_hasta = rs_HC!horhorahasta
                                    Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                                    Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                                End If
                            Else
                                'Si entro despues de las 13 o es Domingo convierto todo a 100%(destino)
                                hora_desde = rs_HC!horhoradesde
                                hora_hasta = rs_HC!horhorahasta
                                Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                                Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                            End If
                            
                            rs_HC.MoveNext
                        Loop
                        
                        'cuando las horas son autorizadas ==> puede que las hs no autorizadas sean mas que las autorizadas ==> debo topear
                        If Total100 > objRsCFG!adcanthoras Then
                            Total100 = objRsCFG!adcanthoras
                        End If
                        If Total50 > objRsCFG!adcanthoras Then
                            Total50 = objRsCFG!adcanthoras
                        End If
                        
                        'Actualizo las hs destino
                            
                            If Total100 <> 0 Then
                                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                OpenRecordset StrSql, objRsAD100
                                
                                If Not objRsAD100.EOF Then
                                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total100, 3)
                                    StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora100 & "," & Round(Total100, 3) & "," & _
                                             CInt(False) & "," & CInt(True) & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                            End If
                            If Total50 <> 0 Then
                                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora50 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                OpenRecordset StrSql, objRsAD100
                                If Not objRsAD100.EOF Then
                                    StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(Total50, 3)
                                    StrSql = StrSql & " WHERE thnro = " & TipoHora50 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,adcanthoras,admanual,advalido) " & _
                                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora50 & "," & Round(Total50, 3) & "," & _
                                             CInt(False) & "," & CInt(True) & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                            End If
                            
                            'Actualizar origen update y delete origen
                            If (Total100 + Total50) = objRsCFG!adcanthoras Then
                             'Borro original
                                 StrSql = " DELETE From gti_acumdiario "
                                 StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                         
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                'Update original de objRsCFG!adcanthoras  -total100
                                StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(objRsCFG!adcanthoras - (Total100 + Total50), 3)
                                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                objConn.Execute StrSql, , adExecuteNoRecords
                            End If
                   
                End If 'If Weekday(p_fecha) = vbSaturday Then
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio Nuevo..."
                End If
                
                
                
                
            End If
        End If
    End If
    '*********************************************************************************************************************************************************************
    '*********************************************************************************************************************************************************************
        
        
  objRsCFG.MoveNext
Loop

FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * 1) & "AD_07 - Fin"
    End If
    
'cierro todo
If objRsCFG.State = adStateOpen Then objRsCFG.Close
If objRsAD.State = adStateOpen Then objRsAD.Close
If objRsAD100.State = adStateOpen Then objRsAD100.Close
If objrhest.State = adStateOpen Then objrhest.Close
If rs_HC.State = adStateOpen Then rs_HC.Close
If rs_AD.State = adStateOpen Then rs_AD.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Cab.State = adStateOpen Then rs_Cab.Close

Set objRsCFG = Nothing
Set objRsAD = Nothing
Set objRsAD100 = Nothing
Set objrhest = Nothing
Set rs_HC = Nothing
Set rs_AD = Nothing
Set rs_Estructura = Nothing
Set rs_Cab = Nothing
End Sub


Public Sub AD_07(p_ternro As Long, p_fecha As Date)
'  --------------------------------------------------------------------------------------------------
'  Archivo:   gtiad07.p
'  Descripción: Convierte los Tipo de Horas despues de autorizar que tengan programas de conversion
'  Autor : ??
'  Fecha : ??
'  Ultima Mod: FGZ - 14/03/2007 ----- Agregado de log
'  --------------------------------------------------------------------------------------------------
'  Ult Modif: CAT - 17/05/2006 Conversion DIVINO SA
'  Ult Modif: FGZ - 25/08/2006
'  Ult Modif: FGZ - 22/09/2006 - Customizacion para Moño Azul (ConversionProd)
'  Ult Modif: FGZ - 22/09/2006 - Customizacion para ICI(las encontré en unos fuentes viejos)
'  Ult Modif: FGZ - 14/11/2006 - Customizacion 12 para AGD()
'  Ult Modif: FGZ - 27/11/2006 - Customizaciones 13, 14 y 15 para AGD()
'  Ult Modif: FGZ - 01/02/2007 - Customizaciones 16, 17 y 18 para Gorina()
'  Ult Modif: Diego Rosso - 21/11/2007 - se agrego la Customizacion 19 para Schneider. Completar()
'  Ult Modif: Diego Rosso - 22/01/2008 - se agrego la Customizacion 20 para MultiVoice. SABADODOMINGO MV()
'  Ult Modif: FGZ - 07/09/2009 - se agrego la Customizacion 22 para TELEARTE.
'  --------------------------------------------------------------------------------------------------
'  Programas Validos:
'  ------------------
'   1.  Conversion              : Conversion estandar a Jornada produccion
'   2.  ConversionProd          : Customizacion para Moño Azul
'   3.  SACO1HORA               : Customizacion para Temaiken
'   4.  REDONDEO                :
'   5.  SABADOS SCHERING        : Customizacion para Schering
'   6.  NormalesEstrada         : Customizacion para Estrada
'   7.  H100DIVINO              : Customizacion para Divino
'   8.  ConvNormales            : Customizacion para ICI
'   9.  Conv50%                 : Customizacion para ICI
'   10. Conv100%                : Customizacion para ICI
'   11. Conv200%                : Customizacion para ICI
'   12. Feriados                : Customizacion para AGD
'   13. Feriados_Estr           : Customizacion para AGD
'   14. Feriados_Trabajados     : Customizacion para AGD
'   15. Feriados_Trabajados_SD  : Customizacion para AGD
'   16. HorasDestajo            : Customizacion para Frig. Gorina
'   17. Adicalmuerzo            : Customizacion para Frig. Gorina
'   18. Peficiencia             : Customizacion para Frig. Gorina
'   19. Completar               : Customizacion para Schneider.
'   20. SABADODOMINGO MV        : Customizacion para MultiVoice.
'   21. TOPEMINIMO              : Customizacion para TRILENIUM extensible al estandar.
'   21. TOPEMINIMO_LV           : Customizacion para TRILENIUM extensible al estandar.
'   21. TOPEMINIMO_SD           : Customizacion para TRILENIUM extensible al estandar.
'   22. VALES_SAT               : Customizacion para TELEARTE.
'   23. TURNO_PLUS              : Customizacion para TELEARTE.
'   24. Feriados_MV             : Customizacion para MULTIVOICE.
'  --------------------------------------------------------------------------------------------------
Dim Horas_Oblig As Single
Dim HorasRes As Single
Dim TotHor As Single
Dim Hora_Dest As Single
Dim Hora_Ori As Single
Dim Nro_Dire As Long
Dim Nro_Ccos As Long
Dim Nro_GSeg As Long
Dim RestoDecimal As Single

Dim EntroAntes11 As Boolean
Dim Total50 As Single
Dim Total100 As Single
Dim TotalHoras As Single

Dim Total_Antes13 As Single
Dim Total_Despues13 As Single

Dim TotalNocturnas As Single
Dim Total150 As Single

Dim TipoHora50 As Long
Dim TipoHora100 As Long
Dim TipoHora150 As Long
Dim TipoNocturna As Integer

Dim TipoHoraNoc100 As Long
Dim TipoHoraNoc150 As Long
Dim TipoHoraFer100 As Long
Dim TipoHoraFer150 As Long

Dim QuedanHs As Boolean
Dim SaldoHS As Single
Dim Dias As Integer
Dim Horas As Integer
Dim Minutos As Integer

Dim Limite1 As String
Dim Limite2 As String

Dim CCosto As Long
Dim Sector As Long
Dim Tenro As Long
Dim Cod_Convenio As String

Dim Tipos_de_Licencias As String
Dim Hay_Licencia As Boolean
Dim Rs_Justif As New ADODB.Recordset
Dim Rs_Lic As New ADODB.Recordset

Dim objRsCFG As New ADODB.Recordset
Dim objRsAD As New ADODB.Recordset
Dim objRsAD100 As New ADODB.Recordset
Dim objrhest As New ADODB.Recordset
Dim rs_HC As New ADODB.Recordset
Dim rs_AD As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Cab As New ADODB.Recordset
Dim SinConvenio As Boolean
Dim ConvenioAnterior As Boolean
Dim rs_TH As New ADODB.Recordset
Dim THOrigen As Long
Dim rs_ST As New ADODB.Recordset
Dim TH_Anormalidad As Long

Dim CantidadDestino As Single
Dim CantidadOrigen As Single
Dim Continua As Boolean

Dim Val_Comida As Single
Dim Val_Merienda As Single
Dim THVal_Comida As Long
Dim THVal_Merienda As Long

Dim THMedioTurno As Long
Dim THTurno As Long
Dim THTurnoyMedio As Long
Dim Val_MedioTurno As Single
Dim Val_Turno As Single
Dim Val_TurnoyMedio As Single

Dim TotHorHHMM As String
Dim Horas_Programadas As Single
Dim Horas_Origen As Single

If depurar Then
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion despues de autorizar. Programas - AD_07() - Inicio"
End If

Nivel_Tab_Log = Nivel_Tab_Log + 1
'21/11/2007 - Diego Rosso
'Agrege  gti_config_horadia.hd_maximo y gti_config_horadia.hd_maximo en el select
'*********************************************************
StrSql = "SELECT gti_config_horadia.hd_thdestino, gti_config_horadia.hd_thorigen, gti_config_horadia.hd_programa, gti_acumdiario.adcanthoras, gti_config_horadia.hd_maximo , gti_config_horadia.hd_minimo "
StrSql = StrSql & " , hd_feriados, hd_nolaborables, hd_laborable "
StrSql = StrSql & " FROM gti_config_horadia "
StrSql = StrSql & " INNER JOIN gti_acumdiario "
StrSql = StrSql & " ON gti_acumdiario.thnro = gti_config_horadia.hd_thorigen "
StrSql = StrSql & " WHERE hd_programa is not null "
StrSql = StrSql & " AND hd_programa <> ''"
StrSql = StrSql & " AND   turnro = " & Nro_Turno & " AND adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
StrSql = StrSql & " ORDER BY hd_nro "
OpenRecordset StrSql, objRsCFG
Do While Not objRsCFG.EOF
    If depurar Then
        Flog.writeline
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Programa: " & objRsCFG!hd_programa
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo Hora Origen: " & objRsCFG!hd_thorigen
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo Hora Destino: " & objRsCFG!hd_thdestino
        Flog.writeline
    End If
    
    If UCase(objRsCFG!hd_programa) = UCase("Conversion") Then
        Call Prog_1_Conversion(p_ternro, p_fecha, objRsCFG!hd_thorigen, objRsCFG!hd_thdestino, objRsCFG!adcanthoras)
    End If
    
    'FGZ - 08/09/2006 - Customizacion para Moño Azul
    If UCase(objRsCFG!hd_programa) = UCase("ConversionProd") Then
        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
        
        'Saco la cantidad de horas del primer dia del turno
        StrSql = "SELECT * FROM gti_turno WHERE turnro = " & Nro_Turno
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!turcanthsprod
        End If
        If Horas_Oblig > 0 Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            If Not objRsAD.EOF Then
                HorasRes = objRsAD!adcanthoras / Horas_Oblig
                
                TotHorHHMM = CHoras(HorasRes, 60)
                
                StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(objRsAD!adcanthoras / Horas_Oblig, 3)
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                TotHorHHMM = CHoras(Round(objRsCFG!adcanthoras / Horas_Oblig, 3), 60)
                
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro, horas,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(objRsCFG!adcanthoras / Horas_Oblig, 3) & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion abortada, Cantidad de horas produccion del turno es 0."
            End If
        End If
    End If
    
    'Esta es una conversión que se aplica en TMK
    If UCase(objRsCFG!hd_programa) = UCase("SACO1HORA") Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
    
            Hora_Dest = objRsAD!adcanthoras

            StrSql = " SELECT estrcodext FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 35 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                Nro_Dire = CLng(objrhest!estrcodext)
            End If
            
            StrSql = " SELECT estrcodext FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 5 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                Nro_Ccos = CLng(objrhest!estrcodext)
            End If
            
            StrSql = " SELECT his_estructura.estrnro FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 7 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                Nro_GSeg = CLng(objrhest!estrnro)
            End If
            
            If InStr(1, "560,525,542,547,543,530,536", CStr(Nro_GSeg)) > 0 Then
                Hora_Dest = Hora_Dest 'Si pertene a algunos  de los grupos de seguridad de la lista, no convertir
            Else
                If Nro_Dire = 47 Then '/* Gerencia de gastronomia */
                    If Hora_Dest >= 6 Then
                        Hora_Dest = Hora_Dest - 0.5
                    Else
                        Hora_Dest = Hora_Dest
                    End If
                Else
                    If Nro_Ccos <> 286 Then  ' /* Centro atencion a la visita */
                        If Hora_Dest >= 6 Then
                            Hora_Dest = Hora_Dest - 1
                        Else
                            Hora_Dest = Hora_Dest
                        End If
                    End If
                End If
            End If
            
            If Not objRsAD.EOF Then
            
                TotHorHHMM = CHoras(Hora_Dest, 60)
                
                StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Hora_Dest
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                TotHorHHMM = CHoras(Hora_Dest, 60)
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Hora_Dest & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
    End If 'esta es una conversion que se aplica en ICI

    'FGZ - 23/09/2004
    If UCase(objRsCFG!hd_programa) = "REDONDEO" Then
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        
        If Not objRsAD.EOF Then
            RestoDecimal = objRsAD!adcanthoras - Fix(objRsAD!adcanthoras)
            
            If RestoDecimal <= 0.25 Then
                HorasRes = Fix(objRsAD!adcanthoras)
            Else
                If RestoDecimal >= 0.251 And RestoDecimal <= 0.75 Then
                    HorasRes = Fix(objRsAD!adcanthoras) + 0.5
                Else
                    HorasRes = Fix(objRsAD!adcanthoras) + 1
                End If
            End If
        
            TotHorHHMM = CHoras(HorasRes, 60)
            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(HorasRes, 3)
            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            TotHorHHMM = CHoras(HorasRes, 60)
            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                     CInt(False) & "," & CInt(True) & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If

    'FGZ - 28/06/2005
    If UCase(objRsCFG!hd_programa) = "SABADOS SCHERING" Then
        If Weekday(p_fecha) = vbSaturday Then
        
            TipoHora50 = 1
            TipoHora100 = 2
        
            StrSql = " SELECT * FROM gti_horcumplido "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
            StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
            StrSql = StrSql & " Order BY thnro, hornro"
            If rs_HC.State = adStateOpen Then rs_HC.Close
            OpenRecordset StrSql, rs_HC
            
            If Not rs_HC.EOF Then
                rs_HC.MoveFirst
                
                If CInt(Mid(rs_HC!horhoradesde, 1, 2)) < 11 Then
                    EntroAntes11 = True
                Else
                    EntroAntes11 = False
                End If
            Else
                'esto no se deberia dar
                If depurar Then
                    Flog.writeline "No se encontraron horas"
                End If
            End If
            
            If EntroAntes11 Then
                '==> de 00:00 a 13:00 son al 50%
                '  y de 13:00 a 24:00 son al 100%
           
                StrSql = " SELECT * FROM gti_horcumplido "
                StrSql = StrSql & " WHERE ternro = " & p_ternro
                StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                StrSql = StrSql & " ORDER BY thnro, hornro"
                If rs_HC.State = adStateOpen Then rs_HC.Close
                OpenRecordset StrSql, rs_HC
                
                Total50 = 0
                Total100 = 0
                Do While Not rs_HC.EOF
                    If CInt(Mid(rs_HC!horhoradesde, 1, 2)) <= 13 Then
                        hora_desde = rs_HC!horhoradesde
                        If CInt(Mid(rs_HC!horhorahasta, 1, 2)) <= 13 Then
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total50 = Total50 + (Dias * 24) + (Horas + (Minutos / 60))
                        Else
                            'hora_hasta = "1259"
                            hora_hasta = "1300"
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total50 = Total50 + (Dias * 24) + (Horas + (Minutos / 60))
                            
                            hora_desde = "1300"
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                        End If
                    Else
                        hora_desde = rs_HC!horhoradesde
                        hora_hasta = rs_HC!horhorahasta
                        Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                        Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                    End If
                    
                    rs_HC.MoveNext
                Loop
                        
                'Actualizo las hs destino
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If Not objRsAD.EOF Then
                    TotHorHHMM = CHoras(Total50, 60)
                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(Total50, 3)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    If Total100 <> 0 Then
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD100
                        
                        If Not objRsAD100.EOF Then
                            TotHorHHMM = CHoras(Total100, 60)
                            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(Total100, 3)
                            StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            TotHorHHMM = CHoras(Total100, 60)
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora100 & "," & TotHorHHMM & "," & Round(Total100, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else
                    TotHorHHMM = CHoras(HorasRes, 60)
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            
            Else
                'se pagan las primeras 5 hs al 50 y si quedan se pagan al 100
                'Total50 = IIf(objRsAD!adcanthoras > 5, 5, objRsAD!adcanthoras)
                'Total100 = IIf(objRsAD!adcanthoras - 5 > 0, objRsAD!adcanthoras - 5, 0)
                
                StrSql = " SELECT * FROM gti_horcumplido "
                StrSql = StrSql & " WHERE ternro = " & p_ternro
                StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND thnro = " & objRsCFG!hd_thorigen
                StrSql = StrSql & " ORDER BY hornro"
                If rs_HC.State = adStateOpen Then rs_HC.Close
                OpenRecordset StrSql, rs_HC
                
                TotalHoras = 0
                Do While Not rs_HC.EOF
                    TotalHoras = TotalHoras + rs_HC!horcant
                
                    rs_HC.MoveNext
                Loop
                Total50 = IIf(TotalHoras > 5, 5, TotalHoras)
                Total100 = IIf(TotalHoras - 5 > 0, TotalHoras - 5, 0)
                If Total100 <> 0 Then
                    QuedanHs = True
                Else
                    QuedanHs = False
                End If
                
                'Actualizo las hs destino
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If Not objRsAD.EOF Then
                    TotHorHHMM = CHoras(Total50, 60)
                    
                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(Total50, 3)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    If Total100 <> 0 Then
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD100
                        
                        If Not objRsAD100.EOF Then
                            TotHorHHMM = CHoras(Total100, 60)
                            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(Total100, 3)
                            StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            TotHorHHMM = CHoras(Total100, 60)
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora100 & "," & TotHorHHMM & "," & Round(Total100, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else
                    TotHorHHMM = CHoras(HorasRes, 60)
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
    End If
    
    'FGZ - 27/07/2005
    'Nueva Politica de Convrersion para Estrada
    If objRsCFG!hd_programa = "NormalesEstrada" Then
        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        StrSql = StrSql & " ORDER BY diaorden"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!diacanthoras
        End If
        TotHorHHMM = CHoras(Horas_Oblig, 60)
        
        StrSql = " SELECT * FROM gti_acumdiario"
        StrSql = StrSql & " INNER JOIN gti_config_tur_hor ON gti_config_tur_hor.thnro = gti_acumdiario.thnro "
        StrSql = StrSql & " WHERE gti_config_tur_hor.turnro = " & Nro_Turno
        StrSql = StrSql & " AND gti_acumdiario.adfecha = " & ConvFecha(p_fecha)
        StrSql = StrSql & " AND gti_acumdiario.ternro = " & p_ternro
        StrSql = StrSql & " AND gti_config_tur_hor.conhornro IN (2,4,5,19)"
        If rs_AD.State = adStateOpen Then rs_AD.Close
        OpenRecordset StrSql, rs_AD
        
        If rs_AD.EOF Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            
            If Not objRsAD.EOF Then
                StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Horas_Oblig
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Horas_Oblig & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        End If
    End If
    'Nueva Politica de Convrersion para Estrada

    'CAT - 14/02/2006
    If UCase(objRsCFG!hd_programa) = "H100DIVINO" Then
            TipoHora100 = 6
            TipoHora150 = 38
            TipoNocturna = 8

            TipoHoraNoc100 = 40
            TipoHoraNoc150 = 41
            TipoHoraFer100 = 42
            TipoHoraFer150 = 43
            
            StrSql = " SELECT * FROM gti_acumdiario "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
            StrSql = StrSql & " AND thnro = " & TipoHora100
            If rs_AD.State = adStateOpen Then rs_AD.Close
            OpenRecordset StrSql, rs_AD
            If rs_AD.EOF Then
                'No hay horas al 100 -> Nada que hacer
                If depurar Then
                    Flog.writeline "No se encontraron Acumulado de Horas. SQL " & StrSql
                End If
            Else
                Total100 = rs_AD!adcanthoras
                If depurar Then
                    Flog.writeline "Total horas 100: " & Total100
                End If

                'Calculo la cantidad de horas Obligatorias
                StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
                OpenRecordset StrSql, objRs
                If Not objRs.EOF Then
                    Horas_Oblig = objRs!diacanthoras
                End If
    
                If Horas_Oblig < 8 Then
                    'si es frando Horas_Oblig = 0 -> entonces vale 8
                    'si es menor que 8, tiene que ser la diferencia de horas
                    'para llegar a 8
                    Horas_Oblig = 8 - Horas_Oblig
                End If
                
                If depurar Then
                    Flog.writeline "Total Obligatorias: " & Horas_Oblig
                End If
                
                If Total100 > Horas_Oblig Then
                    'si las extras 100 son mas que las obligatorias
                    'convertir el exedente en 150
                    
                    TotHorHHMM = CHoras(Horas_Oblig, 60)
                    
                    'Actualizo las horas al 100 como la cantidad de horas Obligatorias
                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(Horas_Oblig, 3)
                    StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    HorasRes = Total100 - Horas_Oblig
                    
                    'Inserto las horas 150 como el exedente de las horas Obligatorias
                    TotHorHHMM = CHoras(HorasRes, 60)
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora150 & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                'Feriado
                If esFeriado Then
                    'reemplazo las Hs100 por Feriado100 y
                    '              Hs150 por Feriado150
                    StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraFer100
                    StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                
                    StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraFer150
                    StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                
                  'Nocturnas
                  Total150 = HorasRes
                  'TotalNocturnas = ContarHs(P_ternro, p_fecha, "2200", p_fecha + 1, "0600")
                  
                  TotalNocturnas = 0
                  StrSql = " SELECT * FROM gti_acumdiario "
                  StrSql = StrSql & " WHERE ternro = " & p_ternro
                  StrSql = StrSql & " AND adfecha = " & ConvFecha(p_fecha)
                  StrSql = StrSql & " AND thnro = " & TipoNocturna
                  If rs_AD.State = adStateOpen Then rs_AD.Close
                  OpenRecordset StrSql, rs_AD
                
                  TotalNocturnas = rs_AD!adcanthoras
                  
                  Do While (HorasRes > 0 Or Total100 > 0) And (TotalNocturnas > 0)
                        If depurar Then
                            Flog.writeline "Total de Horas Nocturnas: " & TotalNocturnas
                            Flog.writeline "Total de Horas 100: " & Total100
                            Flog.writeline "Total de Horas 150: " & Total150
                        End If
                      If Total150 > 0 Then
                          'cambio las 150 por nocturnas 150
                          If TotalNocturnas < Total150 Then
                              TotHorHHMM = CHoras(rs_AD!adcanthoras - TotalNocturnas, 60)
                          
                              StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = adcanthoras - " & Round(TotalNocturnas, 3)
                              StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Actualizo Horas 150. " & StrSql
                              End If
                          
                              TotHorHHMM = CHoras(TotalNocturnas, 60)
                              StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                       " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHoraNoc150 & "," & TotHorHHMM & "," & Round(TotalNocturnas, 3) & "," & _
                                       CInt(False) & "," & CInt(True) & ")"
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Inserto Horas Noct 150. " & StrSql
                              End If
                              Total150 = Total150 - TotalNocturnas
                              TotalNocturnas = 0
                          Else
                              If TotalNocturnas = Total150 Then
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc150
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 150 a Noct 150. " & StrSql
                                  End If
                                  Total150 = 0
                                  TotalNocturnas = 0
                              Else    'TotalNocturnas > Total150
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc150
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora150 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 150 a Noct 150. " & StrSql
                                  End If
                                  
                                  TotalNocturnas = TotalNocturnas - Total150
                                  Total150 = 0
                              End If
                          End If
                      Else
                          'cambio las 100 por nocturnas 100
                          If TotalNocturnas < Total100 Then
                              TotHorHHMM = CHoras(rs_AD!adcanthoras - TotalNocturnas, 60)
                                
                              StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = adcanthoras - " & Round(TotalNocturnas, 3)
                              StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Actualizo Horas 100. " & StrSql
                              End If
                              TotHorHHMM = CHoras(TotalNocturnas, 60)
                              StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                       " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TotHorHHMM & "," & TipoHoraNoc100 & "," & Round(TotalNocturnas, 3) & "," & _
                                       CInt(False) & "," & CInt(True) & ")"
                              objConn.Execute StrSql, , adExecuteNoRecords
                              If depurar Then
                                Flog.writeline "Insert Horas Noct 100. " & StrSql
                              End If
                          
                              Total100 = Total100 - TotalNocturnas
                              TotalNocturnas = 0
                          Else
                              If TotalNocturnas = Total100 Then
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc100
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 100 a Noct 100. " & StrSql
                                  End If
                                  Total100 = 0
                                  TotalNocturnas = 0
                              Else    'TotalNocturnas > Total100
                                  StrSql = " UPDATE gti_acumdiario SET thnro = " & TipoHoraNoc100
                                  StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                  objConn.Execute StrSql, , adExecuteNoRecords
                                  If depurar Then
                                    Flog.writeline "Actualizo Horas 100 a Noct 100. " & StrSql
                                  End If
                                  TotalNocturnas = TotalNocturnas - Total100
                                  Total100 = 0
                              End If
                          End If
                      End If
                  Loop
               End If
            End If
    End If

    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "ConvNormales" Then
        If Weekday(p_fecha) = 1 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If (rs_HC!horhoradesde > rs_HC!horhorahasta) Then
                    objFechasHoras.RestaHs p_fecha, "0000", p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                    HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
                    TotHorHHMM = CHoras(HorasRes, 60)
                    StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                    StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            
            End If
        End If
        If Weekday(p_fecha) = 7 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If rs_HC!horhoradesde >= Limite1 Then
                
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If rs_HC!horhorahasta > Limite1 Then
                        objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, Limite1, Tdias, Thoras, Tmin
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        TotHorHHMM = CHoras(HorasRes, 60)
                        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                    Else
                        objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        TotHorHHMM = CHoras(HorasRes, 60)
                        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If
    
    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "Conv50%" Then
        If Weekday(p_fecha) <> 7 Then
             StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
             StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If Weekday(p_fecha) = 7 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If rs_HC!horhoradesde >= Limite2 Then
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If rs_HC!horhoradesde > Limite1 Then
                        If rs_HC!horhorahasta >= Limite2 Then
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, Limite2, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    Else
                        If rs_HC!horhorahasta >= Limite2 Then
                            objFechasHoras.RestaHs p_fecha, Limite1, p_fecha, Limite2, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            objFechasHoras.RestaHs p_fecha, Limite1, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                            HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                            
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                            StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                            StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "Conv200%" Then
        If Weekday(p_fecha) <> 1 Then
             StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
             StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
    End If
    
    
    
    'FGZ - 22/09/2006 - Estas conversiones las encontre en fuentes viejos (en teoria son para ICI)
    If objRsCFG!hd_programa = "Conv100%" Then
        If Weekday(p_fecha) <> 7 And Weekday(p_fecha) <> 1 Then
             StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
             StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
             objConn.Execute StrSql, , adExecuteNoRecords
        End If
        If (Weekday(p_fecha) = 1) Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If (rs_HC!horhoradesde > rs_HC!horhorahasta) Then
                    objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, "2400", Tdias, Thoras, Tmin
                    HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
                    TotHorHHMM = CHoras(HorasRes, 60)
                    StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                    StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
        
        If Weekday(p_fecha) = 7 Then
            StrSql = " SELECT * FROM gti_horcumplido WHERE "
            StrSql = StrSql & " gti_horcumplido.ternro = " & p_ternro & " AND "
            StrSql = StrSql & " horfecrep = " & ConvFecha(p_fecha) & " AND "
            StrSql = StrSql & " gti_horcumplido.thnro = " & objRsCFG!hd_thorigen
            OpenRecordset StrSql, rs_HC
            If Not rs_HC.EOF Then
                If rs_HC!horhorahasta <= Limite2 And (rs_HC!horhoradesde < rs_HC!horhorahasta) Then
                    StrSql = "DELETE FROM gti_acumdiario WHERE gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " AND"
                    StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    If rs_HC!horhoradesde >= Limite2 Then
                        If (rs_HC!horhoradesde < rs_HC!horhorahasta) Then
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        Else
                            objFechasHoras.RestaHs p_fecha, rs_HC!horhoradesde, p_fecha + 1, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        End If
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        TotHorHHMM = CHoras(HorasRes, 60)
                        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        objFechasHoras.RestaHs p_fecha, Limite2, p_fecha, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                        HorasRes = (Tdias * 24) + (Thoras + (Tmin / 60))
                        
                        TotHorHHMM = CHoras(HorasRes, 60)
                        StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & HorasRes
                        StrSql = StrSql & " Where gti_acumdiario.thnro = " & objRsCFG!hd_thdestino & " And "
                        StrSql = StrSql & " adfecha = " & ConvFecha(p_fecha) & " AND ternro = " & p_ternro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If

    'FGZ - 27/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados") Then
        Set objFeriado.Conexion = objConn
        Set objFeriado.ConexionTraza = CnTraza
        esFeriado = objFeriado.Feriado(p_fecha, Empleado.Ternro, depurar)
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Feriado? " & esFeriado
        End If
        
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        If Not objRsAD.EOF Then
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            If esFeriado Then
                'Reviso que no tenga licencia de:
                Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                
                'Busco si el dia tiene justificacion
                StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                OpenRecordset StrSql, Rs_Justif
                If Not Rs_Justif.EOF Then
                    'Busco la licencia
                    StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                    StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                    StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                    StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                    StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                    'FGZ - 19/05/2010 ------------ Control FT -------------
                    StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                    'FGZ - 19/05/2010 ------------ Control FT -------------
                    OpenRecordset StrSql, Rs_Lic
                    If Rs_Lic.EOF Then
                        Hay_Licencia = False
                    Else
                        Hay_Licencia = True
                    End If
                Else
                    Hay_Licencia = False
                End If
        
                If Not Hay_Licencia Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Hay Licencia. Se paga."
                    End If
                
                    'Si Origen = destino ==> Quedan como estan
                    '                    Sino Creo el destino
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'Inserto el tipo de hora
                        TotHorHHMM = CHoras(Horas_Oblig, 60)
                        
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Horas_Oblig & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'Quedan como estan
                    End If
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hay Licencia. No se paga."
                    End If
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'No se pagan
                        StrSql = " UPDATE gti_acumdiario SET horas = '00:00',adcanthoras = 0"
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'No se pagan ==> borro
                        StrSql = " DELETE gti_acumdiario "
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        Else    'El tipo de hora destino no existe
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El tipo de hora destino no existe"
            End If
            
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            If esFeriado Then
                'Reviso que no tenga licencia de:
                Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                
                'Busco si el dia tiene justificacion
                StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                OpenRecordset StrSql, Rs_Justif
                If Not Rs_Justif.EOF Then
                    'Busco la licencia
                    StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                    StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                    StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                    StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                    StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                    'FGZ - 19/05/2010 ------------ Control FT -------------
                    StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                    'FGZ - 19/05/2010 ------------ Control FT -------------
                    OpenRecordset StrSql, Rs_Lic
                    If Rs_Lic.EOF Then
                        Hay_Licencia = False
                    Else
                        Hay_Licencia = True
                    End If
                Else
                    Hay_Licencia = False
                End If
        
                If Not Hay_Licencia Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No Hay Licencia. Se paga."
                    End If
                    
                    'Si Origen = destino ==> Quedan como estan
                    '                    Sino Creo el destino
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'Inserto el tipo de hora
                        TotHorHHMM = CHoras(Horas_Oblig, 60)
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Horas_Oblig & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'Quedan como estan
                    End If
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hay Licencia. No se paga."
                    End If
                    
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'no las creo porque no se pagan
                    Else
                        'No se pagan ==> las borro
                        StrSql = " DELETE gti_acumdiario"
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para AGD


    'FGZ - 14/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_Estr") Then
        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
        OpenRecordset StrSql, objRsAD
        If Not objRsAD.EOF Then
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            Set objFeriado.Conexion = objConn
            Set objFeriado.ConexionTraza = CnTraza
            esFeriado = objFeriado.Feriado(p_fecha, Empleado.Ternro, depurar)
            
            If esFeriado Then
                If Not Feriado_Por_Estructura Then
                    'Reviso que no tenga licencia de:
                    Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                    
                    'Busco si el dia tiene justificacion
                    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                    OpenRecordset StrSql, Rs_Justif
                    If Not Rs_Justif.EOF Then
                        'Busco la licencia
                        StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                        StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                        StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                        StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                        'FGZ - 19/05/2010 ------------ Control FT -------------
                        StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                        'FGZ - 19/05/2010 ------------ Control FT -------------
                        OpenRecordset StrSql, Rs_Lic
                        If Rs_Lic.EOF Then
                            Hay_Licencia = False
                        Else
                            Hay_Licencia = True
                        End If
                    Else
                        Hay_Licencia = False
                    End If
            
                    If Not Hay_Licencia Then
                        'Si Origen = destino ==> Quedan como estan
                        '                    Sino Creo el destino
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'Inserto el tipo de hora
                            TotHorHHMM = CHoras(Horas_Oblig, 60)
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Horas_Oblig & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'Quedan como estan
                        End If
                    Else
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'No se pagan
                            StrSql = " UPDATE gti_acumdiario SET horas = '00:00', adcanthoras = 0"
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'No se pagan ==> borro
                            StrSql = " DELETE gti_acumdiario "
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else    'Feriado por estructura
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'No se pagan
                        StrSql = " UPDATE gti_acumdiario SET horas = '00:00', adcanthoras = 0"
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'No se pagan ==> borro
                        StrSql = " DELETE gti_acumdiario "
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        Else    'El tipo de hora destino no existe
            'Verifico las condiciones
            
            'Feriado Nacional
            'Quien no trabaje el día feriado, cobra Horas de Feriado. En ese caso siempre son 8 horas.
            '(Salvo si tiene Licencia por Accidente, Vacaciones, Casamiento, Nacimiento, Fallecimiento, Enfermedad, que se verían incluidas en estos conceptos)
            'En el caso que las trabaje, serán las 8 horas de feriado - fijas - más las horas trabajadas.
            
            If esFeriado Then
                If Not Feriado_Por_Estructura Then
                    'Reviso que no tenga licencia de:
                    Tipos_de_Licencias = "2,3,4,5,8,9,13,14"
                    
                    'Busco si el dia tiene justificacion
                    StrSql = "SELECT gti_justificacion.*,gti_tipojust.thnro FROM gti_justificacion INNER JOIN gti_tipojust ON gti_justificacion.tjusnro = gti_tipojust.tjusnro WHERE gti_justificacion.jusnro = " & Nro_Justif
                    OpenRecordset StrSql, Rs_Justif
                    If Not Rs_Justif.EOF Then
                        'Busco la licencia
                        StrSql = "SELECT emp_lic.*,tipdia.thnro,tipdia.tdnro FROM emp_lic "
                        StrSql = StrSql & " INNER JOIN tipdia ON emp_lic.tdnro = tipdia.tdnro AND tipdia.tdnro IN (" & Tipos_de_Licencias & ")"
                        StrSql = StrSql & " WHERE (empleado = " & p_ternro & ") "
                        StrSql = StrSql & " AND (emp_licnro = " & Rs_Justif!juscodext & ")"
                        StrSql = StrSql & " AND (emp_lic.licestnro = 2)" 'Autorizada
                        'FGZ - 19/05/2010 ------------ Control FT -------------
                        StrSql = StrSql & " AND (emp_lic.ft = 0 OR (emp_lic.ft = -1 AND emp_lic.ftap = -1))"
                        'FGZ - 19/05/2010 ------------ Control FT -------------
                        OpenRecordset StrSql, Rs_Lic
                        If Rs_Lic.EOF Then
                            Hay_Licencia = False
                        Else
                            Hay_Licencia = True
                        End If
                    Else
                        Hay_Licencia = False
                    End If
            
                    If Not Hay_Licencia Then
                        'Si Origen = destino ==> Quedan como estan
                        '                    Sino Creo el destino
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'Inserto el tipo de hora
                            TotHorHHMM = CHoras(Horas_Oblig, 60)
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Horas_Oblig & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'Quedan como estan
                        End If
                    Else
                        If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                            'no las creo porque no se pagan
                        Else
                            'No se pagan ==> las borro
                            StrSql = " DELETE gti_acumdiario"
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                Else    'Feriado por estructura
                    If objRsCFG!hd_thdestino <> objRsCFG!hd_thorigen Then
                        'No se pagan
                    Else
                        'No se pagan
                    End If
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para AGD

    'FGZ - 27/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_Trabajados") Then
        'las obligatorias serian las de horigen
        'Las excedentes serian las destino
    
        'Calculo la cantidad de horas Obligatorias
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!diacanthoras
        End If
        If Horas_Oblig < 8 Then
            'si es frando Horas_Oblig = 0 -> entonces vale 8
            'si es menor que 8, tiene que ser la diferencia de horas
            'para llegar a 8
            Horas_Oblig = 8 - Horas_Oblig
        End If
        If depurar Then
            Flog.writeline "Total Obligatorias: " & Horas_Oblig
        End If
        TotHor = objRsCFG!adcanthoras
        If depurar Then
            Flog.writeline "Cant Hs Origen: " & TotHor
        End If
        If TotHor > Horas_Oblig Then
            'Horas Origen
            TotHorHHMM = CHoras(Horas_Oblig, 60)
            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Horas_Oblig
            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            objConn.Execute StrSql, , adExecuteNoRecords
            
            'Horas excedentes
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            If objRsAD.EOF Then
                'Inserto
                TotHorHHMM = CHoras(TotHor - Horas_Oblig, 60)
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & (TotHor - Horas_Oblig) & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else    'El tipo de hora destino no existe
                'Actualizo el valor
                TotHorHHMM = CHoras(TotHor - Horas_Oblig, 60)
                StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = " & (TotHor - Horas_Oblig)
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
        End If
    End If
    'Nueva Politica de Convrersion para AGD

    'FGZ - 27/11/2006
    'Nueva Politica de Convrersion para AGD
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_Trabajados_SD") Then
        'las obligatorias serian las de horigen
        'Las excedentes serian las destino
        If Weekday(p_fecha) > 1 And Weekday(p_fecha) < 7 Then 'Dia de semana
            'Calculo la cantidad de horas Obligatorias
            StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
            OpenRecordset StrSql, objRs
            If Not objRs.EOF Then
                Horas_Oblig = objRs!diacanthoras
            End If
            If Horas_Oblig < 8 Then
                'si es frando Horas_Oblig = 0 -> entonces vale 8
                'si es menor que 8, tiene que ser la diferencia de horas
                'para llegar a 8
                Horas_Oblig = 8 - Horas_Oblig
            End If
            If depurar Then
                Flog.writeline "Total Obligatorias: " & Horas_Oblig
            End If
            TotHor = objRsCFG!adcanthoras
            If depurar Then
                Flog.writeline "Cant Hs Origen: " & TotHor
            End If
            If TotHor > Horas_Oblig Then
                'Horas Origen
                TotHorHHMM = CHoras(Horas_Oblig, 60)
                StrSql = " UPDATE gti_acumdiario SET horas=" & TotHorHHMM & ",adcanthoras = " & Horas_Oblig
                StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
                'Horas excedentes
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If objRsAD.EOF Then
                    'Inserto
                    TotHorHHMM = CHoras(TotHor - Horas_Oblig, 60)
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & (TotHor - Horas_Oblig) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else    'El tipo de hora destino no existe
                    'Actualizo el valor
                    TotHorHHMM = CHoras(TotHor - Horas_Oblig, 60)
                    StrSql = " UPDATE gti_acumdiario SET horas=" & TotHorHHMM & ",adcanthoras = " & (TotHor - Horas_Oblig)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else
                'Quedan como estan
            End If
        Else
            If Weekday(p_fecha) = 1 Then 'Domingo
                'Se convierten todas las horas al destino
                TotHor = objRsCFG!adcanthoras
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If objRsAD.EOF Then
                    'Inserto
                    TotHorHHMM = CHoras(TotHor, 60)
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & TotHor & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else    'El tipo de hora destino no existe
                    'Actualizo el valor
                    TotHorHHMM = CHoras(TotHor, 60)
                    StrSql = " UPDATE gti_acumdiario SET horas=" & TotHorHHMM & ",adcanthoras = " & TotHor
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            Else    'Sabado
            
                StrSql = " SELECT * FROM gti_horcumplido "
                StrSql = StrSql & " WHERE ternro = " & p_ternro
                StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                StrSql = StrSql & " AND thnro = " & objRsCFG!hd_thorigen
                StrSql = StrSql & " ORDER BY hornro"
                If rs_HC.State = adStateOpen Then rs_HC.Close
                OpenRecordset StrSql, rs_HC
                
                Total_Antes13 = 0
                Total_Despues13 = 0
                Do While Not rs_HC.EOF
                    If CInt(Mid(rs_HC!horhoradesde, 1, 2)) <= 13 Then
                        hora_desde = rs_HC!horhoradesde
                        If CInt(Mid(rs_HC!horhorahasta, 1, 2)) <= 13 Then
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total_Antes13 = Total_Antes13 + (Dias * 24) + (Horas + (Minutos / 60))
                        Else
                            hora_hasta = "1300"
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total_Antes13 = Total_Antes13 + (Dias * 24) + (Horas + (Minutos / 60))
                            
                            hora_desde = "1300"
                            hora_hasta = rs_HC!horhorahasta
                            Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                            Total_Despues13 = Total_Despues13 + (Dias * 24) + (Horas + (Minutos / 60))
                        End If
                    Else
                        hora_desde = rs_HC!horhoradesde
                        hora_hasta = rs_HC!horhorahasta
                        Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                        Total_Despues13 = Total_Despues13 + (Dias * 24) + (Horas + (Minutos / 60))
                    End If
                    
                    rs_HC.MoveNext
                Loop
                
                'Actualizo las hs Origen
                If Total_Antes13 <> 0 Then
                    TotHorHHMM = CHoras(Total_Antes13, 60)
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = " & Round(Total_Antes13, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thorigen & "," & TotHorHHMM & "," & Round(Total_Antes13, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        'llevo la cantidad a 0
                        StrSql = " UPDATE gti_acumdiario SET horas = '00:00', adcanthoras = " & Round(0, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        'No hago nada
                    End If
                End If
                
                If Total_Despues13 <> 0 Then
                    TotHorHHMM = CHoras(Total_Despues13, 60)
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = " & Round(Total_Despues13, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(Total_Despues13, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para AGD

    'FGZ - 01/02/2007
    'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
    If UCase(objRsCFG!hd_programa) = UCase("HORASDESTAJO") Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HORASDESTAJO -->"
            End If
            'Busco Centro de Costo y el Sector
            CCosto = 0
            Sector = 0
            
            Tenro = 5   'Centro De Costo
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                CCosto = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Centro de Costo."
                End If
            End If
        
            Tenro = 2   'Sector
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Sector = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Sector"
                End If
            End If
        
            If CCosto <> 0 And Sector <> 0 Then
                StrSql = " SELECT ordtcant FROM orden_trabajo "
                StrSql = StrSql & " WHERE estrnro2 = " & Sector
                StrSql = StrSql & " AND (estrnro3 = 0 OR estrnro3 is null OR estrnro3 = " & CCosto & ")"
                StrSql = StrSql & " AND ordtfecdesde = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, rs_Cab
                If Not rs_Cab.EOF Then
                    HorasRes = rs_Cab!ordtcant
                    
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        TotHorHHMM = CHoras(HorasRes, 60)
                        
                        StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(HorasRes, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        TotHorHHMM = CHoras(HorasRes, 60)
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                Else
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No hay ordenes para la fecha: " & StrSql
                    End If
                End If
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "CCosto = 0 OR Sector = 0. " & CCosto & " y " & Sector
                End If
            End If
        End If
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        
    'FGZ - 01/02/2007
    'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
    If UCase(objRsCFG!hd_programa) = UCase("ADICALMUERZO") Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "ADICALMUERZO -->"
            End If
            'Busco Centro de Costo y el Sector
            CCosto = 0
            Sector = 0
            
            Tenro = 5   'Centro De Costo
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                CCosto = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Centro de Costo."
                End If
            End If
        
            Tenro = 2   'Sector
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Sector = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Sector"
                End If
            End If
        
            If CCosto <> 0 And Sector <> 0 Then
                StrSql = " SELECT ordtcant FROM orden_trabajo "
                StrSql = StrSql & " WHERE estrnro2 = " & Sector
                StrSql = StrSql & " AND (estrnro3 = 0 OR estrnro3 is null OR estrnro3 = " & CCosto & ")"
                StrSql = StrSql & " AND ordtfecdesde = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, rs_Cab
                If Not rs_Cab.EOF Then
                    HorasRes = rs_Cab!ordtcant
                    
                    If HorasRes > 400 Then
                        HorasRes = 40
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD
                        If Not objRsAD.EOF Then
                            TotHorHHMM = CHoras(HorasRes, 60)
                        
                            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(HorasRes, 3)
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                End If
            End If
        End If
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        
        'FGZ - 01/02/2007
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        If UCase(objRsCFG!hd_programa) = UCase("PEFICIENCIA") Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "PEFICIENCIA -->"
            End If
            'Busco Centro de Costo y el Sector
            CCosto = 0
            Sector = 0
            
            Tenro = 5   'Centro De Costo
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                CCosto = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Centro de Costo."
                End If
            End If
        
            Tenro = 2   'Sector
            StrSql = " SELECT estrnro FROM his_estructura "
            StrSql = StrSql & " WHERE ternro = " & p_ternro
            StrSql = StrSql & " AND tenro =" & Tenro
            StrSql = StrSql & " AND (htetdesde <= " & ConvFecha(p_fecha) & ")"
            StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= htethasta) or (htethasta is null))"
            OpenRecordset StrSql, rs_Estructura
            If Not rs_Estructura.EOF Then
                Sector = rs_Estructura!estrnro
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Sector"
                End If
            End If
        
            If CCosto <> 0 And Sector <> 0 Then
                StrSql = " SELECT ordtcant FROM orden_trabajo "
                StrSql = StrSql & " WHERE estrnro2 = " & Sector
                StrSql = StrSql & " AND (estrnro3 = 0 OR estrnro3 is null OR estrnro3 = " & CCosto & ")"
                StrSql = StrSql & " AND ordtfecdesde = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, rs_Cab
                If Not rs_Cab.EOF Then
                    HorasRes = rs_Cab!ordtcant
                    TotHorHHMM = CHoras(HorasRes, 60)
                    
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(HorasRes, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                End If
            End If
        End If
        'Nueva Convrersion para Frigorifico Gorina ---------------------------------------
        
    'Diego Rosso - 21/11/2007
    'CAS-04990- Nueva Conversión GTI - Schneider Electric -----------------------
    'Completar la cantidad de Hs. faltantes para llegar a totalizar las horas mínimas de jornada
    'en los dias Configurados.
    'Tiene en cuenta el mínimo de Horas Normales realizadas para que esto tenga efecto.
    If UCase(objRsCFG!hd_programa) = UCase("Completar") Then
        'Llamo a la politica
        Call Politica(571)
        TH_Anormalidad = st_TipoHora1
        
        If depurar Then
           Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Tipo de Hora de Anormalidad " & TH_Anormalidad
           Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Lista de Dias " & st_ListaTH
           Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Dia de la semama a procesar " & Weekday(p_fecha)
           StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
           OpenRecordset StrSql, objRsAD
           If Not objRsAD.EOF Then
                If depurar Then
                   Flog.writeline Espacios(Tabulador * 4) & "Antes de la conversion hay " & objRsAD!adcanthoras & " horas destino"
                End If
           Else
                If depurar Then
                   Flog.writeline Espacios(Tabulador * 4) & "Antes de la conversion NO hay horas destino"
                End If
           End If
        End If
        
        'Chequear que este configurada la politica
        If Not EsNulo(st_ListaTH) Then
            'weekday. Chequeo que el dia que estoy procesando este en la lista sino no lo proceso
            If InStr(1, st_ListaTH, Weekday(p_fecha)) > 0 Then
                    'Chequeo que las horas que hizo el empleado esten dentro del minimo y maximo configurado para la conversion
                    If (objRsCFG!adcanthoras <= objRsCFG!hd_maximo) And (objRsCFG!adcanthoras >= objRsCFG!hd_minimo) Then
        
                         'Busca Cantidad de Horas de Día para el turno del Empleado
                         StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
                         OpenRecordset StrSql, objRs
                         If Not objRs.EOF Then
                             Horas_Oblig = objRs!diacanthoras ' Cantidad de horas configuradas para el dia
                         Else
                            Horas_Oblig = 0
                            If depurar Then
                               Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontró la Cantidad de Horas del día " & Nro_Dia
                            End If
                         End If
                         
                         'Cantidad de horas a convertir
                         HorasRes = Horas_Oblig - objRsCFG!adcanthoras
                         If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas Obligatorias: " & Horas_Oblig
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas de origen encontradas: " & objRsCFG!adcanthoras
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas a compensar: " & HorasRes
                         End If
                         
                         'chequea negativo
                         If HorasRes > 0 Then
                            'Revisar si existe ST por esa cantidad de hs
                            '   si existe ==> borro anormalidad y genero compensacion
                            '   sino ... nada
                            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TH_Anormalidad & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            StrSql = StrSql & " AND adcanthoras = " & HorasRes
                            OpenRecordset StrSql, rs_ST
                            If Not rs_ST.EOF Then
                                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                OpenRecordset StrSql, objRsAD
                                If Not objRsAD.EOF Then
                                    If depurar Then
                                       Flog.writeline Espacios(Tabulador * 4) & "Actualizo horas destino --> " & Round(objRsAD!adcanthoras + HorasRes, 3)
                                    End If
                                    
                                    TotHorHHMM = CHoras(HorasRes, 60)
                                    
                                    'StrSql = " UPDATE gti_acumdiario SET adcanthoras = " & Round(objRsAD!adcanthoras + HorasRes, 3)
                                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(HorasRes, 3)
                                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    If depurar Then
                                       Flog.writeline Espacios(Tabulador * 4) & "Inserto horas destino --> " & Round(HorasRes, 3)
                                    End If
                                    
                                    TotHorHHMM = CHoras(HorasRes, 60)
                                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(HorasRes, 3) & "," & _
                                             CInt(False) & "," & CInt(True) & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                                
                                'Borro el tipo de hora del AD
                                StrSql = " DELETE gti_acumdiario WHERE thnro = " & TH_Anormalidad & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                StrSql = StrSql & " AND adcanthoras = " & HorasRes
                                objConn.Execute StrSql, , adExecuteNoRecords
                                
                                'Cambio el la condicion de anormalidad del HC anormalidad
                                StrSql = "UPDATE gti_horcumplido set normnro = 10, normnro2 = 10 WHERE thnro = " & TH_Anormalidad & " AND ternro = " & p_ternro & " AND horfecrep = " & ConvFecha(p_fecha)
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                If depurar Then
                                   Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No compensa porque no encontró la anormalidad. "
                                End If
                            End If
                         Else 'HorasRes > 0
                            If depurar Then
                               Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La cantidad de horas de que registradas en el dia iguala o supera la cantidad de horas configuradas. "
                            End If
                         End If 'HorasRes > 0
                        
                    Else '(objRsCFG!diacanthoras <= objRsCFG!hd_maximo) And (objRsCFG!diacanthoras >= objRsCFG!hd_minimo)
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La cantidad de horas de origen no entran entre el minimo y maximo configurado."
                        End If
                    End If '(objRsCFG!diacanthoras <= objRsCFG!hd_maximo) And (objRsCFG!diacanthoras >= objRsCFG!hd_minimo)
            Else 'InStr(1, st_ListaTH, Weekday(p_fecha)) > 0
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El dia no esta dentro de la lista de dias configurada."
                End If
            End If 'InStr(1, st_ListaTH, Weekday(p_fecha)) > 0
        Else '(Not st_ListaTH Is Null) Or (st_ListaTH <> "")
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "La politica 571 no esta configurada."
            End If
        End If '(Not st_ListaTH Is Null) Or (st_ListaTH <> "")
    End If
        
        
    '*********************************************************************************************************************************************************************
    'Diego Rosso - 22/01/2008
    'Sabado hasta las 13 hs  genere Hs adicionales Aut. y después del sabado a las 13 y hasta el domingo genera Hs 100%  según corresponda.
    '*********************************************************************************************************************************************************************
    If UCase(objRsCFG!hd_programa) = "SABADODOMINGO MV" Then
        Tenro = 55   'Convenio 'Tenro = 19   'Convenio
        SinConvenio = True
        StrSql = " SELECT estructura.estrnro, estructura.estrcodext FROM his_estructura "
        StrSql = StrSql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro"
        StrSql = StrSql & " WHERE his_estructura.ternro = " & p_ternro
        StrSql = StrSql & " AND estructura.tenro =" & Tenro
        StrSql = StrSql & " AND (his_estructura.htetdesde <= " & ConvFecha(p_fecha) & ")"
        StrSql = StrSql & " AND ((" & ConvFecha(p_fecha) & " <= his_estructura.htethasta) or (his_estructura.htethasta is null))"
        OpenRecordset StrSql, rs_Estructura
        If Not rs_Estructura.EOF Then
            If Not EsNulo(rs_Estructura!estrcodext) Then
                ConvenioAnterior = IIf(rs_Estructura!estrcodext = "0", True, False)
                SinConvenio = False
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "El Convenio " & rs_Estructura!estrnro & " no tiene configurado el codigo externo. No se ejecutará la coversión."
                End If
            End If
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "No se encontro la esructura Convenio. No se ejecutará la coversión."
            End If
        End If
        
        If Not SinConvenio Then
            If ConvenioAnterior Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio Anterior..."
                End If
                If Weekday(p_fecha) = vbSaturday Or Weekday(p_fecha) = vbSunday Then
                    TipoHora100 = objRsCFG!hd_thdestino
                    THOrigen = objRsCFG!hd_thorigen
                    
                    'Busco el horario trabajado en el dia
                    StrSql = " SELECT * FROM gti_horcumplido "
                    StrSql = StrSql & " WHERE ternro = " & p_ternro
                    StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                    StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                    StrSql = StrSql & " AND thnro = " & THOrigen
                    StrSql = StrSql & " Order BY thnro, hornro"
                    OpenRecordset StrSql, rs_HC
                    If rs_HC.EOF Then
                        'quiere decir que este tipo de hora fué autorizado ==>
                        'debo buscar el tipo de hora no autorizado y
                        'topearlo a la cantidad de hs autorizadas
                        StrSql = "SELECT thnro FROM tiphora WHERE thautpor = " & THOrigen
                        OpenRecordset StrSql, rs_TH
                        If Not rs_TH.EOF Then
                            THOrigen = rs_TH!thnro

                            StrSql = " SELECT * FROM gti_horcumplido "
                            StrSql = StrSql & " WHERE ternro = " & p_ternro
                            StrSql = StrSql & " AND hordesde = " & ConvFecha(p_fecha)
                            StrSql = StrSql & " AND horhasta = " & ConvFecha(p_fecha)
                            StrSql = StrSql & " AND thnro = " & THOrigen
                            StrSql = StrSql & " Order BY thnro, hornro"
                            OpenRecordset StrSql, rs_HC
                        End If
                    End If
                    
                      If depurar Then
                            Flog.writeline "Fecha a procesar:" & p_fecha & " Correspondiente a un dia Sabado"
                      End If
                        'Si es sabado
                        '==> de 00:00 a 13:00 quedan igual
                        '  y de 13:00 a 24:00 son al 100%
                        'Si es Domingo
                        'Todas las horas se pasan al 100%
                        Total100 = 0
                        Do While Not rs_HC.EOF
                            
                            If CInt(Mid(rs_HC!horhoradesde, 1, 4)) <= 1300 And Weekday(p_fecha) = vbSaturday Then
        '                        hora_desde = rs_HC!horhoradesde
                                If CInt(Mid(rs_HC!horhorahasta, 1, 2)) >= 13 Then
                                    hora_desde = "1300"
                                    hora_hasta = rs_HC!horhorahasta
                                    Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                                    Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                                End If
                            Else
                                'Si entro despues de las 13 o es Domingo convierto todo a 100%(destino)
                                hora_desde = rs_HC!horhoradesde
                                hora_hasta = rs_HC!horhorahasta
                                Call Restar_Horas(p_fecha, hora_desde, p_fecha, hora_hasta, Dias, Horas, Minutos)
                                Total100 = Total100 + (Dias * 24) + (Horas + (Minutos / 60))
                            End If
                            
                            rs_HC.MoveNext
                        Loop
                        
                        'cuando las horas son autorizadas ==> puede que las hs no autorizadas sean mas que las autorizadas ==> debo topear
                        If Total100 > objRsCFG!adcanthoras Then
                            Total100 = objRsCFG!adcanthoras
                        End If
                        'Actualizo las hs destino
                            If Total100 <> 0 Then
                                TotHorHHMM = CHoras(Total100, 60)
                                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                OpenRecordset StrSql, objRsAD100
                                If Not objRsAD100.EOF Then
                                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(Total100, 3)
                                    StrSql = StrSql & " WHERE thnro = " & TipoHora100 & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & TipoHora100 & "," & TotHorHHMM & "," & Round(Total100, 3) & "," & _
                                             CInt(False) & "," & CInt(True) & ")"
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                            
                            
                                'Actualizar origen update y delete origen
                                If Total100 = objRsCFG!adcanthoras Then
                                 'Borro original
                                     StrSql = " DELETE From gti_acumdiario "
                                     StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    'Update original de objRsCFG!adcanthoras  -total100
                                    TotHorHHMM = CHoras(objRsCFG!adcanthoras - Total100, 60)
                                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(objRsCFG!adcanthoras - Total100, 3)
                                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                End If
                            End If
                End If 'If Weekday(p_fecha) = vbSaturday Then
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio Nuevo..."
                End If
            End If
        End If
    End If
    
    
    'FGZ - 12/12/2007 - Tope de Hs
    'Deja la cantidad minima como Origen y la diferencia entre la cantidad y el minimo como destino
    If UCase(objRsCFG!hd_programa) = UCase("TopeMinimo") Then
        Continua = False
        If esFeriado Then
            If objRsCFG!hd_feriados Then
                Continua = True
            End If
        Else
            If Dia_Libre Then
                If objRsCFG!hd_nolaborables Then
                    Continua = True
                End If
            Else
                If objRsCFG!hd_laborable Then
                    Continua = True
                End If
            End If
        End If
    
        If Continua Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convirtiendo...."
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs: " & objRsCFG!adcanthoras
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  Minimo      : " & objRsCFG!hd_minimo
            End If
            If objRsCFG!adcanthoras > objRsCFG!hd_minimo Then
                CantidadDestino = objRsCFG!adcanthoras - objRsCFG!hd_minimo
                CantidadOrigen = objRsCFG!hd_minimo
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Se Convierte en"
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs Origen: " & CantidadOrigen
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs Destino: " & CantidadDestino
                End If
                'Busco las de origen
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HS Origen.."
                End If
                
                TotHorHHMM = CHoras(CantidadOrigen, 60)
                
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If Not objRsAD.EOF Then
                    StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(CantidadOrigen)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thorigen & "," & TotHorHHMM & "," & Round(CantidadOrigen, 3) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                
                'Busco las de origen destino
                If CantidadDestino <> 0 Then
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HS destino.."
                    End If
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                        End If
                        
                        TotHorHHMM = CHoras(objRsAD!adcanthoras + CantidadDestino, 60)
                        
                        StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(CantidadDestino, 3)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                        End If
                    Else
                        TotHorHHMM = CHoras(CantidadDestino, 60)
                        
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(CantidadDestino, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'FGZ - 16/12/2007 - Tope de Hs
    'Deja la cantidad minima como Origen y la diferencia entre la cantidad y el minimo como destino
    If UCase(objRsCFG!hd_programa) = UCase("TopeMinimo_LV") Then
        Continua = False
        If esFeriado Then
            If objRsCFG!hd_feriados Then
                Continua = True
            End If
        Else
            If Dia_Libre Then
                If objRsCFG!hd_nolaborables Then
                    Continua = True
                End If
            Else
                If objRsCFG!hd_laborable Then
                    Continua = True
                End If
            End If
        End If
        If Continua Then
            Select Case Weekday(p_fecha)
            Case 2, 3, 4, 5, 6: ' Lunes a viernes ==> se hace la conversion
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convirtiendo...."
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs: " & objRsCFG!adcanthoras
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  Minimo      : " & objRsCFG!hd_minimo
                End If
                                    
                If objRsCFG!adcanthoras > objRsCFG!hd_minimo Then
                    CantidadDestino = objRsCFG!adcanthoras - objRsCFG!hd_minimo
                    CantidadOrigen = objRsCFG!hd_minimo
                        
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Se Convierte en"
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs Origen: " & CantidadOrigen
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs Destino: " & CantidadDestino
                    End If
                    'Busco las de origen
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HS Origen.."
                    End If
                    TotHorHHMM = CHoras(CantidadDestino, 60)
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(CantidadOrigen)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thorigen & "," & TotHorHHMM & "," & Round(CantidadOrigen, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    
                    'Busco las de origen destino
                    If CantidadDestino <> 0 Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HS destino.."
                        End If
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD
                        If Not objRsAD.EOF Then
                            If depurar Then
                                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                            End If
                            
                            TotHorHHMM = CHoras(objRsAD!adcanthoras + CantidadDestino, 60)
                            
                            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(CantidadDestino, 3)
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                            If depurar Then
                                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                            End If
                        Else
                            TotHorHHMM = CHoras(CantidadDestino, 60)
                            
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(CantidadDestino, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            If depurar Then
                                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                            End If
                        End If
                    End If
                End If
            Case 1, 7:
                'los sabados y domingos no se hace la conversion
            End Select
        End If
    End If
    
    'FGZ - 16/12/2007 - Tope de Hs
    'Deja la cantidad minima como Origen y la diferencia entre la cantidad y el minimo como destino
    If UCase(objRsCFG!hd_programa) = UCase("TopeMinimo_SD") Then
        Continua = False
        If esFeriado Then
            If objRsCFG!hd_feriados Then
                Continua = True
            End If
        Else
            If Dia_Libre Then
                If objRsCFG!hd_nolaborables Then
                    Continua = True
                End If
            Else
                If objRsCFG!hd_laborable Then
                    Continua = True
                End If
            End If
        End If
        If Continua Then
        
            Select Case Weekday(p_fecha)
            Case 1, 7: ' Sabados y Domingo ==> se hace la conversion
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convirtiendo...."
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs: " & objRsCFG!adcanthoras
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  Minimo      : " & objRsCFG!hd_minimo
                End If
                                    
                If objRsCFG!adcanthoras > objRsCFG!hd_minimo Then
                    CantidadDestino = objRsCFG!adcanthoras - objRsCFG!hd_minimo
                    CantidadOrigen = objRsCFG!hd_minimo
                        
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Se Convierte en"
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs Origen: " & CantidadOrigen
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Cantidad de Hs Destino: " & CantidadDestino
                    End If
                        
                    'Busco las de origen
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HS Origen.."
                    End If
                    
                    TotHorHHMM = CHoras(CantidadOrigen, 60)
                    
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & Round(CantidadOrigen)
                        StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                    Else
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thorigen & "," & TotHorHHMM & "," & Round(CantidadOrigen, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    
                    'Busco las de origen destino
                    If CantidadDestino <> 0 Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "HS destino.."
                        End If
                        StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        OpenRecordset StrSql, objRsAD
                        If Not objRsAD.EOF Then
                            If depurar Then
                                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                            End If
                            
                            TotHorHHMM = CHoras(objRsAD!adcanthoras + CantidadDestino, 60)
                            
                            StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(CantidadDestino, 3)
                            StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                            objConn.Execute StrSql, , adExecuteNoRecords
                            If depurar Then
                                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                            End If
                        Else
                            TotHorHHMM = CHoras(CantidadDestino, 60)
                            StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                     " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & Round(CantidadDestino, 3) & "," & _
                                     CInt(False) & "," & CInt(True) & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                            If depurar Then
                                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                            End If
                        End If
                    End If
                End If
            Case 2, 3, 4, 5, 6:
                'de Lunes a Viernes no se hace la conversion
            End Select
        End If
    End If
    
    'Esta es una conversión que se aplica en TELEARTE para el convenio SAT
    If UCase(objRsCFG!hd_programa) = UCase("VALES_SAT") Then
            'Tipos de hora destino
            THVal_Comida = 42
            THVal_Merienda = 43
            
            StrSql = " SELECT estrcodext FROM his_estructura, estructura "
            StrSql = StrSql & " WHERE his_estructura.tenro = 19 and htethasta is null and ternro = " & p_ternro & " and "
            StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
            OpenRecordset StrSql, objrhest
            If Not objrhest.EOF Then
                If Not EsNulo(objrhest!estrcodext) Then
                    Cod_Convenio = objrhest!estrcodext
                Else
                    Cod_Convenio = ""
                    If depurar Then
                        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio sin Codigo Externo. No aplica. "
                    End If
                End If
            Else
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Empleado sin convenio. No aplica. "
                End If
            End If
            If UCase(Cod_Convenio) = "SAT" Then
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Aplica para el Convenio. "
                End If
            
                'revisar condiciones
                'a)De 1 hs a 2.5 horas extra generdas= 1 vale merienda
                'b)De 3 hs a 5.5 horas extra generadas= 1 vale comida
                'c)De 6 hs a 6.5 horas extra generadas= 1 vale comida y 1 vale merienda
                'd)De 7 hs a 8.5 horas extras generadas = 2 vales comida
                'e)De 9 hs en adelante horas extras generadas = 2 vales comida y 1 vale merienda.
                
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If objRsAD.EOF Then
                    Hora_Ori = 0
                Else
                    Hora_Ori = objRsAD!adcanthoras
                End If
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas extras " & Hora_Ori
                End If
                
                Val_Comida = 0
                Val_Merienda = 0
                
                Select Case Hora_Ori
                Case Is >= 9
                    Val_Comida = 2
                    Val_Merienda = 1
                Case Is >= 7
                    Val_Comida = 2
                    Val_Merienda = 0
                Case Is >= 6
                    Val_Comida = 1
                    Val_Merienda = 1
                Case Is >= 3
                    Val_Comida = 1
                    Val_Merienda = 0
                Case Is >= 1
                    Val_Comida = 0
                    Val_Merienda = 1
                Case Else
                    Val_Comida = 0
                    Val_Merienda = 0
                End Select
                If depurar Then
                    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Corresponden. " & Val_Comida & " Vales Comida y " & Val_Merienda & " Vales Medienda."
                End If
                'Vales Comida
                If Val_Comida > 0 Then
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & THVal_Comida & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                        End If
                        TotHorHHMM = CHoras(objRsAD!adcanthoras + Val_Comida, 60)
                        StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(Val_Comida, 3)
                        StrSql = StrSql & " WHERE thnro = " & THVal_Comida & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                        End If
                    Else
                        TotHorHHMM = CHoras(Val_Comida, 60)
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & THVal_Comida & "," & TotHorHHMM & "," & Round(Val_Comida, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                        End If
                    End If
                End If
                
                'Vales Merienda
                If Val_Merienda > 0 Then
                    StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & THVal_Merienda & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    OpenRecordset StrSql, objRsAD
                    If Not objRsAD.EOF Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                        End If
                        
                        TotHorHHMM = CHoras(objRsAD!adcanthoras + Val_Merienda, 60)
                        
                        StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(Val_Merienda, 3)
                        StrSql = StrSql & " WHERE thnro = " & THVal_Merienda & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                        End If
                    Else
                        TotHorHHMM = CHoras(Val_Merienda, 60)
                        StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                                 " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & THVal_Merienda & "," & TotHorHHMM & "," & Round(Val_Merienda, 3) & "," & _
                                 CInt(False) & "," & CInt(True) & ")"
                        objConn.Execute StrSql, , adExecuteNoRecords
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                        End If
                    End If
                End If
            End If
    End If
    
    'Esta es una conversión que se aplica en TELEARTE
    If UCase(objRsCFG!hd_programa) = UCase("TURNO_PLUS") Then
        Call Prog_23_Turno_Plus(p_ternro, p_fecha, objRsCFG!hd_thorigen, objRsCFG!hd_thdestino, objRsCFG!adcanthoras)
    End If
    
    
    'FGZ - 10/09/2010
    'Nueva Politica de Convrersion para MV ------------------------------------------------
        'Si las Hs Origen < Hs Teóricas ==>
        '   hs Destino < 1
        '       Ejemplo le programaron 5 horas y vino 4 horas Ò Hs Destino = 0.8
        'SINO
        '    SI HS Origen > Hs Teóricas ==>
        '       HS Destino > 1 ==> HS Destino = 1 (luego del truncado)
        '           Ejemplos
        '               Le programaron 6, trabajó 7 horas y no le autorizaron la hora extra
        '               Le programaron 6, trabajó 7 horas y le autorizaron la hora extra
        '
        '
        '        No necesito controlar las horas extras, si están autorizadas o no.
        '        Si las horas de de origen son mayores que la teóricas implica en principio que hubo extras.
        '        SI las extras estuvieron autorizadas ==> HS destino = 1 y si las horas extras no estuvieran autorizadas Hs Destino serán > 1, que luego del truncado quedan en 1
        'SINO
        '    Las Hs Origen son iguales a Hs Teóricas y por ende
        '       Hs Destino = 1
    
        'OBS
        '   Si las hs teoricas o programadas = 0 ==> Hs Destino = 0
    
    
    If UCase(objRsCFG!hd_programa) = UCase("Feriados_MV") Then
        'las Programadas serian las de teoricas para el dia
        '
    
        'Calculo la cantidad de horas Programadas
        StrSql = "SELECT * FROM gti_dias WHERE dianro = " & Nro_Dia
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Programadas = objRs!diacanthoras
        End If
        If Horas_Programadas = 0 Then
            If depurar Then
                Flog.writeline "Horas Programadas = 0 "
            End If
        Else
            If depurar Then
                Flog.writeline "Horas Programadas: " & Horas_Programadas
            End If
            Horas_Origen = objRsCFG!adcanthoras
        
            If Horas_Origen >= Horas_Programadas Then
                TotHor = 1
            Else
                TotHor = Round(Horas_Origen / Horas_Programadas, 2)
            End If
            
            If TotHor > 0 Then
                'Horas destino
                StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                OpenRecordset StrSql, objRsAD
                If objRsAD.EOF Then
                    'Inserto
                    TotHorHHMM = CHoras(TotHor, 60)
                    StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro,horas,adcanthoras,admanual,advalido) " & _
                             " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & objRsCFG!hd_thdestino & "," & TotHorHHMM & "," & (TotHor) & "," & _
                             CInt(False) & "," & CInt(True) & ")"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else    'El tipo de hora destino no existe
                    'Actualizo el valor
                    TotHorHHMM = CHoras(TotHor, 60)
                    StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = " & (TotHor)
                    StrSql = StrSql & " WHERE thnro = " & objRsCFG!hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
            End If
        End If
    End If
    'Nueva Politica de Convrersion para MV
    
    
    
    '*********************************************************************************************************************************************************************
  objRsCFG.MoveNext
Loop

FIN:
    Nivel_Tab_Log = Nivel_Tab_Log - 1
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion despues de autorizar. Programas - AD_07() - FIN"
    End If
    
'cierro todo
If objRsCFG.State = adStateOpen Then objRsCFG.Close
If objRsAD.State = adStateOpen Then objRsAD.Close
If objRsAD100.State = adStateOpen Then objRsAD100.Close
If objrhest.State = adStateOpen Then objrhest.Close
If rs_HC.State = adStateOpen Then rs_HC.Close
If rs_AD.State = adStateOpen Then rs_AD.Close
If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
If rs_Cab.State = adStateOpen Then rs_Cab.Close

Set objRsCFG = Nothing
Set objRsAD = Nothing
Set objRsAD100 = Nothing
Set objrhest = Nothing
Set rs_HC = Nothing
Set rs_AD = Nothing
Set rs_Estructura = Nothing
Set rs_Cab = Nothing
End Sub


Public Sub CrearAD(P_NroTer As Long)
Dim Acum As Double
Dim Acum_HS As String
Dim Aux_HS As String
Dim AuxThnro As Integer

Dim objrsAC_D As New ADODB.Recordset
Dim objrsWFAd As New ADODB.Recordset

Acum = 0
AuxThnro = -1

'StrSql = "SELECT thnro, cant_hs FROM " & TTempWFAd & " GROUP BY thnro ASC"
StrSql = "SELECT * FROM " & TTempWFAd
'StrSql = StrSql & " WHERE acumula = 0"
StrSql = StrSql & " ORDER BY thnro"
OpenRecordset StrSql, objrsWFAd

Do While Not objrsWFAd.EOF
    ' esto es para inicializar acum cuando cambia el thnro
    If objrsWFAd!thnro <> AuxThnro Then
        Acum = 0
        Acum_HS = "00:00"
    End If
    
    Acum = Acum + objrsWFAd!Cant_hs
    Call SHoras(Acum_HS, IIf(IsNull(objrsWFAd!Horas), "00:00", objrsWFAd!Horas), Acum_HS)
    
    AuxThnro = objrsWFAd!thnro
    
    StrSql = "SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(p_fecha) & _
        " AND ternro = " & P_NroTer & " AND thnro = " & objrsWFAd!thnro
    OpenRecordset StrSql, objrsAC_D
    
    If objrsAC_D.EOF Then
        'inserto
        StrSql = "INSERT INTO gti_acumdiario (ternro,thnro,adfecha,horas, adcanthoras, admanual) VALUES (" & _
                P_NroTer & "," & objrsWFAd!thnro & "," & ConvFecha(p_fecha) & ",'" & Acum_HS & "'," & Acum & ", 0)"
    Else
        'modifico el campo de cantidad de horas
        'Call SHoras(Acum_HS, objrsWFAd!Horas, Aux_HS)
        
        StrSql = "UPDATE gti_acumdiario SET horas = '" & objrsWFAd!Horas & "', adcanthoras = " & objrsWFAd!Cant_hs & ", advalido = -1 WHERE adfecha = " & ConvFecha(p_fecha) & _
            " AND ternro = " & P_NroTer & " AND thnro = " & objrsWFAd!thnro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Acum = 0
    Acum_HS = "00:00"
    
    ' esto no se porque
    objrsAC_D.Close
    
Siguiente:
    objrsWFAd.MoveNext
Loop
    
    
    
    
End Sub

Public Sub CrearAD2(P_NroTer As Long)
Dim Acum As Single
Dim Acum_HS As String
Dim Aux_HS As String
Dim AuxThnro As Integer

Dim objrsAC_D As New ADODB.Recordset
Dim objrsWFAd As New ADODB.Recordset

Acum = 0
AuxThnro = -1

'StrSql = "SELECT thnro, cant_hs FROM " & TTempWFAd & " GROUP BY thnro ASC"
StrSql = "SELECT * FROM " & TTempWFAd
StrSql = StrSql & " WHERE acumula = 0"
StrSql = StrSql & " ORDER BY thnro"
OpenRecordset StrSql, objrsWFAd

Do While Not objrsWFAd.EOF
    ' esto es para inicializar acum cuando cambia el thnro
    If objrsWFAd!thnro <> AuxThnro Then
        Acum = 0
        Acum_HS = "00:00"
    End If
    
    Acum = Acum + objrsWFAd!Cant_hs
    Call SHoras(Acum_HS, IIf(IsNull(objrsWFAd!Horas), "00:00", objrsWFAd!Horas), Acum_HS)
    
    AuxThnro = objrsWFAd!thnro
    
    StrSql = "SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(p_fecha) & _
        " AND ternro = " & P_NroTer & " AND thnro = " & objrsWFAd!thnro
    OpenRecordset StrSql, objrsAC_D
    
    If objrsAC_D.EOF Then
        'inserto
        StrSql = "INSERT INTO gti_acumdiario (ternro,thnro,adfecha,horas, adcanthoras, admanual) VALUES (" & _
                P_NroTer & "," & objrsWFAd!thnro & "," & ConvFecha(p_fecha) & ",'" & Acum_HS & "'," & Acum & ", 0)"
    Else
        Call SHoras(Acum_HS, objrsWFAd!Horas, Aux_HS)
        'modifico el campo de cantidad de horas
        StrSql = "UPDATE gti_acumdiario SET horas = '" & Aux_HS & "',adcanthoras = adcanthoras + " & objrsWFAd!Cant_hs & ", advalido = -1 WHERE adfecha = " & ConvFecha(p_fecha) & _
            " AND ternro = " & P_NroTer & " AND thnro = " & objrsWFAd!thnro
    End If
    objConn.Execute StrSql, , adExecuteNoRecords
    
    Acum = 0
    Acum_HS = "00:00"
    
    ' esto no se porque
    objrsAC_D.Close
    
Siguiente:
    objrsWFAd.MoveNext
Loop
End Sub

Private Sub Limpia()
    StrSql = "DELETE FROM " & TTempWFAd
    objConn.Execute StrSql, , adExecuteNoRecords
End Sub

Public Sub Autoriza(Fecha As Date, NroTer As Long, Cant As Integer)
'/*----------------------------------------------------------------------------
'Archivo:   gtiautex.p
'  Descripci¢n: Discrimina horas en Autorizadas o No Autorizadas.
'  Autor: Marchese, Juan M.
'  Creado: 25/10/2000
'Modificado: FGZ - 07/05/2009
'               Le agregué el control de la cantidad de hs minimas que deben quedar sin autorizar para que genere la anormalidad de hs sin autorizar
'               La cantidad de minutos se configura con la politica 576
'---------------------------------------------------------------------------*/
Dim HorasRes As Single
Dim autorizada As Integer
Dim no_autorizada As Integer
Dim horasaut As Single
Dim Hora As Integer
Dim Firmado As Boolean
Dim entro As Boolean

Dim rs As New ADODB.Recordset
Dim rsAutdet As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset
Dim rs_FT As New ADODB.Recordset

Dim TotHorHHMM As String
Dim Lista_Horas As String

Nivel_Tab_Log = Nivel_Tab_Log + 1
If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Entró en AUTORIZACION: "
End If

'Seteo de los tipos de horas que, por mas que tengan la marca de autorizable no se autorizan.
Call Politica(577)
If Not EsNulo(ListaNoAutorizable) Then
    Lista_Horas = ListaNoAutorizable
Else
    Lista_Horas = "0"
End If

Cant = 0
entro = False
autorizada = 0
no_autorizada = 0

StrSql = "SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha)
StrSql = StrSql & " AND ternro = " & NroTer
StrSql = StrSql & " AND thnro NOT IN (" & Lista_Horas & ")"
OpenRecordset StrSql, objRs

Do While Not objRs.EOF
    entro = False
    Hora = objRs!thnro
    
    StrSql = "select thautpor,thdesautpor from tiphora where thnro = " & objRs!thnro
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hora Autorizada: " & rs!thautpor
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hora no Autorizada: " & rs!thdesautpor
        End If
        autorizada = rs!thautpor
        no_autorizada = rs!thdesautpor
        
        rs.Close
        
        'Se recorren todos los partes del tercero autorizables
        If (autorizada = 0) And (no_autorizada = 0) Then
             GoTo NextProgress
        Else
            StrSql = "select * from gti_cabparte " & _
            " INNER JOIN gti_autdet on gti_cabparte.gcpnro = gti_autdet.gcpnro " & _
            " WHERE (gcpdesde <= " & ConvFecha(Fecha) & _
            ") and (gcphasta >= " & ConvFecha(Fecha) & ") AND " & _
            " ternro = " & objRs!Ternro & " and thnro = " & objRs!thnro & _
            " and gadautorizable = -1 and " & _
            "((gadfecdesde <= " & ConvFecha(Fecha) & " or (gadfecdesde is null)) and " & _
            "(gadfechasta >= " & ConvFecha(Fecha) & " or (gadfechasta is null)))"
            OpenRecordset StrSql, rsAutdet
            
            Do While Not rsAutdet.EOF
                entro = True
                
                'FGZ - 31/05/2010  --------------------------------------------------------------------------
                'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
                StrSql = "SELECT input_ft.idnro,input_ft.origen, gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
                StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
                StrSql = StrSql & " WHERE idtipoinput = 8 "
                StrSql = StrSql & " AND origen = " & rsAutdet!gcpnro
                OpenRecordset StrSql, rs_FT
                If Not rs_FT.EOF Then
                    'El parte fué cargado fuera de termimo
                    If rs_FT!ftap = -1 Then
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * 6) & "Parte de Horas extras fuera de termino aprobado."
                        End If
                        Firmado = True
                        Call InsertarFT(rs_FT!idnro, 8, rs_FT!Origen)
                    Else
                        If depurar Then
                            Flog.writeline Espacios(Tabulador * 6) & "Parte de Horas extras fuera de termino no aprobado. Se descarta."
                        End If
                        Firmado = False
                    End If
                Else
                    'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para partes de Autorizacion de horas
                    OpenRecordset "select * from cystipo where cystipnro = 1", rs
                    'Verificar si esta en el NIVEL FINAL DE FIRMA
                    If rs!cystipact = -1 Then
                        StrSql = "select * from cysfirmas where cysfirfin = -1 and " & _
                        "cysfircodext = '" & rsAutdet!gcpnro & "' and cystipnro = 1"
                        rs.Close
                        OpenRecordset StrSql, rs
                        If rs.EOF Then
                            Firmado = False
                        Else
                            Firmado = True
                        End If
                    Else
                        Firmado = True
                    End If
                    rs.Close
                End If
                
                '/* Si no est  firmado y es autorizable, desautoriza las horas */
                If Not Firmado And (rsAutdet!gadautorizable = -1) Then
                    StrSql = "UPDATE gti_acumdiario SET thnro = " & no_autorizada & _
                    " WHERE ternro = " & objRs!Ternro & _
                    " AND adfecha = " & ConvFecha(objRs!adfecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 07/05/2009 --------------------------------------------------------------
                    'Control sobre la cantidad minima de hs sin autorizar que deben quedar para que se genere la anormalidad
                    If HayMinimoExtrasSinAutorizar Then
                        'Reviso cual es la cantidad de hs que quedarian sin autorizar
                        StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                        StrSql = StrSql & " WHERE ternro = " & objRs!Ternro
                        StrSql = StrSql & " AND adfecha = " & ConvFecha(objRs!adfecha)
                        StrSql = StrSql & " AND thnro = " & objRs!thnro
                        OpenRecordset StrSql, rsAD
                        If Not rsAD.EOF Then
                            If rsAD!adcanthoras >= MinimoExtrasSinAutorizar Then
                                'C.A.T Si no esta firmado cargo la anormalidad de Extras no Autorizadas
                                StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                                " WHERE ternro = " & objRs!Ternro & _
                                " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                                " AND thnro = " & objRs!thnro
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                'No genero la anormalidad
                            End If
                        Else
                            'el registro no está ==> no hago nada (esto no debiera suceder)
                        End If
                    Else
                        'Quiere decir que no está configurada la politica o no tiene alcance
                        'POr cuestiones de compatibilidad lo dejo como esaba en el estandar antes de esta modificacion
                        
                        'C.A.T Si no esta firmado cargo la anormalidad de Extras no Autorizadas
                        StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                        " WHERE ternro = " & objRs!Ternro & _
                        " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                        " AND thnro = " & objRs!thnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    'FGZ - 07/05/2009 --------------------------------------------------------------
                    
                    GoTo NextProgress
                End If
                
                '/* Si alguno de los tipos de horas no est  configurado, pasa a otro AD */
                
                StrSql = "select * from tiphora where thnro = " & autorizada
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    GoTo NextProgress
                End If
                
                StrSql = "select * from tiphora where thnro = " & no_autorizada
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    GoTo NextProgress
                End If
                
                horasaut = rsAutdet!gadhoras
                
                If objRs!adcanthoras > horasaut Then
                    HorasRes = objRs!adcanthoras - horasaut
                    
                    StrSql = "delete from gti_acumdiario where" & _
                    " ternro = " & objRs!Ternro & _
                    " and adfecha = " & ConvFecha(Fecha) & _
                    " and thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    '/* Horas Autorizadas */
                    
                    StrSql = "select * from gti_acumdiario where" & _
                    " ternro = " & NroTer & _
                    " and thnro = " & autorizada & _
                    " and adfecha = " & ConvFecha(Fecha)
                    OpenRecordset StrSql, rs
                    
                    If horasaut > 0 Then
                        If (rs.EOF) Then
                            
                            TotHorHHMM = CHoras(horasaut, 60)
                            
                            StrSql = "insert into gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
                            " values(" & NroTer & "," & autorizada & "," & _
                            ConvFecha(Fecha) & "," & TotHorHHMM & "," & horasaut & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            TotHorHHMM = CHoras(rs!adcanthoras + horasaut, 60)
                            
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & _
                            rs!adcanthoras + horasaut & " where " & _
                            " ternro = " & objRs!Ternro & _
                            " and adfecha = " & ConvFecha(Fecha) & _
                            " and thnro = " & autorizada
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                    rs.Close
                    
                    '/* Horas no Autorizadas */
                    StrSql = "select * from gti_acumdiario where" & _
                    " ternro = " & NroTer & _
                    " and thnro = " & no_autorizada & _
                    " and adfecha = " & ConvFecha(Fecha)
                    OpenRecordset StrSql, rs
    
                    If HorasRes > 0 Then
                        If (rs.EOF) Then
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = "insert into gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
                            " values(" & NroTer & "," & no_autorizada & "," & _
                            ConvFecha(Fecha) & "," & TotHorHHMM & "," & HorasRes & ")"
                        Else
                            TotHorHHMM = CHoras(rs!adcanthoras + HorasRes, 60)
                            StrSql = "update gti_acumdiario set adcanthoras = " & _
                            rs!adcanthoras + HorasRes & " where " & _
                            " ternro = " & objRs!Ternro & _
                            " and adfecha = " & ConvFecha(Fecha) & _
                            " and thnro = " & no_autorizada
                        End If
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        'FGZ - 07/05/2009 --------------------------------------------------------------
                        'Control sobre la cantidad minima de hs sin autorizar que deben quedar para que se genere la anormalidad
                        If HayMinimoExtrasSinAutorizar Then
                            'Reviso cual es la cantidad de hs que quedarian sin autorizar
                            StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                            StrSql = StrSql & " WHERE ternro = " & objRs!Ternro
                            StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha)
                            StrSql = StrSql & " AND thnro = " & no_autorizada
                            OpenRecordset StrSql, rsAD
                            If Not rsAD.EOF Then
                                If rsAD!adcanthoras >= MinimoExtrasSinAutorizar Then
                                    'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                                    StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                                    " WHERE ternro = " & objRs!Ternro & _
                                    " AND horfecrep = " & ConvFecha(Fecha) & _
                                    " AND thnro = " & no_autorizada
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    'No genero la anormalidad
                                End If
                            Else
                                'el registro no está ==> no hago nada (esto no debiera suceder)
                            End If
                        Else
                            'Quiere decir que no está configurada la politica o no tiene alcance
                            'POr cuestiones de compatibilidad lo dejo como esaba en el estandar antes de esta modificacion
                            
                            'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                            StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                            " WHERE ternro = " & objRs!Ternro & _
                            " AND horfecrep = " & ConvFecha(Fecha) & _
                            " AND thnro = " & no_autorizada
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                        'FGZ - 07/05/2009 --------------------------------------------------------------
                    End If
                    rs.Close
                    
                Else
                    '/* El total de horas es autorizable */
                    horasaut = objRs!adcanthoras
                    
                    StrSql = "DELETE FROM gti_acumdiario WHERE" & _
                    " ternro = " & objRs!Ternro & _
                    " AND adfecha = " & ConvFecha(Fecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    
                    'FGZ 18/09/08 Debo actualizar la hora original por si se le habia generado la onormalidad anteriormente.
                    StrSql = "UPDATE gti_horcumplido SET normnro = 0 " & _
                    " WHERE ternro = " & objRs!Ternro & _
                    " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    
                    StrSql = "select * from gti_acumdiario where" & _
                    " ternro = " & NroTer & _
                    " and thnro = " & autorizada & _
                    " and adfecha = " & ConvFecha(Fecha)
                    OpenRecordset StrSql, rs
                    
                    If (horasaut > 0) Then
                        If (rs.EOF) Then
                            TotHorHHMM = CHoras(horasaut, 60)
                            
                            StrSql = "INSERT INTO gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
                            " values(" & NroTer & "," & autorizada & "," & _
                            ConvFecha(Fecha) & "," & TotHorHHMM & "," & horasaut & ")"
                        Else
                            TotHorHHMM = CHoras(rs!adcanthoras + horasaut, 60)
                            
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ", adcanthoras = " & _
                            rs!adcanthoras + horasaut & " where " & _
                            " ternro = " & objRs!Ternro & _
                            " and adfecha = " & ConvFecha(Fecha) & _
                            " and thnro = " & autorizada
                        End If
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    rs.Close
                End If
                
                rsAutdet.MoveNext
            Loop
            
            If Not entro Then
                'cambiar en acumdiario el tipo de hora
                
                'objRs!thnro = no_autorizada
                StrSql = "update gti_acumdiario set thnro = " & _
                no_autorizada & " WHERE " & _
                " ternro = " & objRs!Ternro & _
                " and adfecha = " & ConvFecha(Fecha) & _
                " and thnro = " & objRs!thnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                
                'FGZ - 07/05/2009 --------------------------------------------------------------
                'Control sobre la cantidad minima de hs sin autorizar que deben quedar para que se genere la anormalidad
                If HayMinimoExtrasSinAutorizar Then
                    'Reviso cual es la cantidad de hs que quedarian sin autorizar
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE ternro = " & objRs!Ternro
                    StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha)
                    StrSql = StrSql & " AND thnro = " & objRs!thnro
                    OpenRecordset StrSql, rsAD
                    If Not rsAD.EOF Then
                        If rsAD!adcanthoras >= MinimoExtrasSinAutorizar Then
                            'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                            StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                            " WHERE ternro = " & objRs!Ternro & _
                            " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                            " AND thnro = " & objRs!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'No genero la anormalidad
                        End If
                    Else
                        'el registro no está ==> no hago nada (esto no debiera suceder)
                    End If
                Else
                    'Quiere decir que no está configurada la politica o no tiene alcance
                    'POr cuestiones de compatibilidad lo dejo como esaba en el estandar antes de esta modificacion
                    
                    'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                    StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                    " WHERE ternro = " & objRs!Ternro & _
                    " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                'FGZ - 07/05/2009 --------------------------------------------------------------
            End If
        End If
    Else
        ' El tipo de Hora no tiene configurado Autoriza y no autoriza
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hora NO tiene configurado Autoriza y No Autoriza: " & rs!thautpor
        End If
        'Exit Sub
    End If
NextProgress:
    objRs.MoveNext
Loop

If objRs.State = adStateOpen Then objRs.Close
If rsAutdet.State = adStateOpen Then rsAutdet.Close
If rs_FT.State = adStateOpen Then rs_FT.Close




'Se recorren todos los partes del tercero directos (crean directamente AD)
StrSql = "select * from gti_cabparte " & _
    " INNER JOIN gti_autdet on gti_cabparte.gcpnro = gti_autdet.gcpnro " & _
    "WHERE (gcpdesde <= " & ConvFecha(Fecha) & _
    ") and (gcphasta >= " & ConvFecha(Fecha) & ")" & _
    " AND ternro = " & NroTer & _
    " and gadautorizable = 0 and " & _
    "((gadfecdesde <= " & ConvFecha(Fecha) & " or (gadfecdesde is null)) and " & _
    "(gadfechasta >= " & ConvFecha(Fecha) & " or (gadfechasta is null)))"
OpenRecordset StrSql, rsAutdet
Do While Not rsAutdet.EOF
    'FGZ - 31/05/2010  --------------------------------------------------------------------------
    'Verifico que no haya sido generado fuera de termino y en ese caso reviso que esté aprobado
    StrSql = "SELECT gti_cabparte.ft, gti_cabparte.ftap FROM input_ft "
    StrSql = StrSql & " INNER JOIN gti_cabparte ON input_ft.origen = gti_cabparte.gcpnro "
    StrSql = StrSql & " WHERE idtipoinput = 8 "
    StrSql = StrSql & " AND origen = " & rsAutdet!gcpnro
    OpenRecordset StrSql, rs_FT
    If Not rs_FT.EOF Then
        'El parte fué cargado fuera de termimo
        If rs_FT!ftap = -1 Then
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Parte de Horas extras fuera de termino aprobado."
            End If
            Firmado = True
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * 6) & "Parte de Horas extras fuera de termino aprobado. Se descarta."
            End If
            Firmado = False
        End If
    Else
        'Verificar si esta en el NIVEL FINAL DE FIRMA ACTIVO para partes de Autorizacion de horas
        OpenRecordset "select * from cystipo where cystipnro = 1", rs
        'Verificar si esta en el NIVEL FINAL DE FIRMA
        If rs!cystipact = -1 Then
            StrSql = "select * from cysfirmas where cysfirfin = -1 and " & _
            "cysfircodext = '" & rsAutdet!gcpnro & "' and cystipnro = 1"
            rs.Close
            OpenRecordset StrSql, rs
            If rs.EOF Then
                Firmado = False
            Else
                Firmado = True
            End If
        Else
            Firmado = True
        End If
        rs.Close
    End If
    
    If Firmado Then
        StrSql = "select * from gti_acumdiario where" & _
        " ternro = " & rsAutdet!Ternro & _
        " and thnro = " & rsAutdet!thnro & _
        " and adfecha = " & ConvFecha(Fecha)
        OpenRecordset StrSql, rs
        If rs.EOF Then
            TotHorHHMM = CHoras(rsAutdet!gadhoras, 60)
            StrSql = "INSERT INTO gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
            " values(" & rsAutdet!Ternro & "," & rsAutdet!thnro & "," & _
            ConvFecha(Fecha) & "," & TotHorHHMM & "," & rsAutdet!gadhoras & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            TotHorHHMM = CHoras(rs!adcanthoras + rsAutdet!gadhoras, 60)
            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & _
            rs!adcanthoras + rsAutdet!gadhoras & " where " & _
            " ternro = " & rsAutdet!Ternro & _
            " and adfecha = " & ConvFecha(Fecha) & _
            " and thnro = " & rsAutdet!thnro
            objConn.Execute StrSql, , adExecuteNoRecords
            Cant = Cant + 1
        End If
        rs.Close
    End If
    rsAutdet.MoveNext
Loop

'libero todo
    If objRs.State = adStateOpen Then objRs.Close
    If rsAutdet.State = adStateOpen Then rsAutdet.Close
    If rs_FT.State = adStateOpen Then rs_FT.Close
End Sub


'antes de las modificaciones por input fuera de termino
Public Sub autoriza_old(Fecha As Date, NroTer As Long, Cant As Integer)
'/*----------------------------------------------------------------------------
'Archivo:   gtiautex.p
'  Descripci¢n: Discrimina horas en Autorizadas o No Autorizadas.
'  Autor: Marchese, Juan M.
'  Creado: 25/10/2000
'Modificado: FGZ - 07/05/2009
'               Le agregué el control de la cantidad de hs minimas que deben quedar sin autorizar para que genere la anormalidad de hs sin autorizar
'               La cantidad de minutos se configura con la politica 576
'---------------------------------------------------------------------------*/
Dim HorasRes As Single
Dim autorizada As Integer
Dim no_autorizada As Integer
Dim horasaut As Single
Dim Hora As Integer
Dim Firmado As Boolean
Dim entro As Boolean

Dim rs As New ADODB.Recordset
Dim rsAutdet As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset

Dim TotHorHHMM As String

Nivel_Tab_Log = Nivel_Tab_Log + 1
If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Entró en AUTORIZACION: "
End If


Cant = 0
entro = False
autorizada = 0
no_autorizada = 0

StrSql = "SELECT * FROM gti_acumdiario WHERE adfecha = " & ConvFecha(Fecha) & _
" AND ternro = " & NroTer
OpenRecordset StrSql, objRs

Do While Not objRs.EOF
    entro = False
    Hora = objRs!thnro
    
    StrSql = "select thautpor,thdesautpor from tiphora where thnro = " & objRs!thnro
    OpenRecordset StrSql, rs
    
    If Not rs.EOF Then
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "hora Autorizada: " & rs!thautpor
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "hora Autorizada: " & rs!thdesautpor
        End If
        autorizada = rs!thautpor
        no_autorizada = rs!thdesautpor
        
        rs.Close
        
        'Se recorren todos los partes del tercero autorizables
        If (autorizada = 0) And (no_autorizada = 0) Then
             GoTo NextProgress
        Else
            StrSql = "select * from gti_cabparte " & _
            " INNER JOIN gti_autdet on gti_cabparte.gcpnro = gti_autdet.gcpnro " & _
            " WHERE (gcpdesde <= " & ConvFecha(Fecha) & _
            ") and (gcphasta >= " & ConvFecha(Fecha) & ") AND " & _
            " ternro = " & objRs!Ternro & " and thnro = " & objRs!thnro & _
            " and gadautorizable = -1 and " & _
            "((gadfecdesde <= " & ConvFecha(Fecha) & " or (gadfecdesde is null)) and " & _
            "(gadfechasta >= " & ConvFecha(Fecha) & " or (gadfechasta is null)))"
            OpenRecordset StrSql, rsAutdet
            
            Do While Not rsAutdet.EOF
                entro = True
                
                OpenRecordset "select * from cystipo where cystipnro = 1", rs
                'Verificar si esta en el NIVEL FINAL DE FIRMA
                If rs!cystipact = -1 Then
                    StrSql = "select * from cysfirmas where cysfirfin = -1 and " & _
                    "cysfircodext = '" & rsAutdet!gcpnro & "' and cystipnro = 1"
                    rs.Close
                    OpenRecordset StrSql, rs
                    If rs.EOF Then
                        Firmado = False
                    Else
                        Firmado = True
                    End If
                Else
                    Firmado = True
                End If
                rs.Close
                
                '/* Si no est  firmado y es autorizable, desautoriza las horas */
                If Not Firmado And (rsAutdet!gadautorizable = -1) Then
                    StrSql = "UPDATE gti_acumdiario SET thnro = " & no_autorizada & _
                    " WHERE ternro = " & objRs!Ternro & _
                    " AND adfecha = " & ConvFecha(objRs!adfecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    'FGZ - 07/05/2009 --------------------------------------------------------------
                    'Control sobre la cantidad minima de hs sin autorizar que deben quedar para que se genere la anormalidad
                    If HayMinimoExtrasSinAutorizar Then
                        'Reviso cual es la cantidad de hs que quedarian sin autorizar
                        StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                        StrSql = StrSql & " WHERE ternro = " & objRs!Ternro
                        StrSql = StrSql & " AND adfecha = " & ConvFecha(objRs!adfecha)
                        StrSql = StrSql & " AND thnro = " & objRs!thnro
                        OpenRecordset StrSql, rsAD
                        If Not rsAD.EOF Then
                            If rsAD!adcanthoras >= MinimoExtrasSinAutorizar Then
                                'C.A.T Si no esta firmado cargo la anormalidad de Extras no Autorizadas
                                StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                                " WHERE ternro = " & objRs!Ternro & _
                                " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                                " AND thnro = " & objRs!thnro
                                objConn.Execute StrSql, , adExecuteNoRecords
                            Else
                                'No genero la anormalidad
                            End If
                        Else
                            'el registro no está ==> no hago nada (esto no debiera suceder)
                        End If
                    Else
                        'Quiere decir que no está configurada la politica o no tiene alcance
                        'POr cuestiones de compatibilidad lo dejo como esaba en el estandar antes de esta modificacion
                        
                        'C.A.T Si no esta firmado cargo la anormalidad de Extras no Autorizadas
                        StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                        " WHERE ternro = " & objRs!Ternro & _
                        " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                        " AND thnro = " & objRs!thnro
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    'FGZ - 07/05/2009 --------------------------------------------------------------
                    
                    GoTo NextProgress
                End If
                
                '/* Si alguno de los tipos de horas no est  configurado, pasa a otro AD */
                
                StrSql = "select * from tiphora where thnro = " & autorizada
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    GoTo NextProgress
                End If
                
                StrSql = "select * from tiphora where thnro = " & no_autorizada
                OpenRecordset StrSql, rs
                If rs.EOF Then
                    GoTo NextProgress
                End If
                
                horasaut = rsAutdet!gadhoras
                
                If objRs!adcanthoras > horasaut Then
                    HorasRes = objRs!adcanthoras - horasaut
                    
                    StrSql = "delete from gti_acumdiario where" & _
                    " ternro = " & objRs!Ternro & _
                    " and adfecha = " & ConvFecha(Fecha) & _
                    " and thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    '/* Horas Autorizadas */
                    
                    StrSql = "select * from gti_acumdiario where" & _
                    " ternro = " & NroTer & _
                    " and thnro = " & autorizada & _
                    " and adfecha = " & ConvFecha(Fecha)
                    OpenRecordset StrSql, rs
                    
                    If horasaut > 0 Then
                        If (rs.EOF) Then
                            
                            TotHorHHMM = CHoras(horasaut, 60)
                            
                            StrSql = "insert into gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
                            " values(" & NroTer & "," & autorizada & "," & _
                            ConvFecha(Fecha) & "," & TotHorHHMM & "," & horasaut & ")"
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            TotHorHHMM = CHoras(rs!adcanthoras + horasaut, 60)
                            
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & _
                            rs!adcanthoras + horasaut & " where " & _
                            " ternro = " & objRs!Ternro & _
                            " and adfecha = " & ConvFecha(Fecha) & _
                            " and thnro = " & autorizada
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                    End If
                    rs.Close
                    
                    '/* Horas no Autorizadas */
                    StrSql = "select * from gti_acumdiario where" & _
                    " ternro = " & NroTer & _
                    " and thnro = " & no_autorizada & _
                    " and adfecha = " & ConvFecha(Fecha)
                    OpenRecordset StrSql, rs
    
                    If HorasRes > 0 Then
                        If (rs.EOF) Then
                            TotHorHHMM = CHoras(HorasRes, 60)
                            StrSql = "insert into gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
                            " values(" & NroTer & "," & no_autorizada & "," & _
                            ConvFecha(Fecha) & "," & TotHorHHMM & "," & HorasRes & ")"
                        Else
                            TotHorHHMM = CHoras(rs!adcanthoras + HorasRes, 60)
                            StrSql = "update gti_acumdiario set adcanthoras = " & _
                            rs!adcanthoras + HorasRes & " where " & _
                            " ternro = " & objRs!Ternro & _
                            " and adfecha = " & ConvFecha(Fecha) & _
                            " and thnro = " & no_autorizada
                        End If
                        objConn.Execute StrSql, , adExecuteNoRecords
                        
                        'FGZ - 07/05/2009 --------------------------------------------------------------
                        'Control sobre la cantidad minima de hs sin autorizar que deben quedar para que se genere la anormalidad
                        If HayMinimoExtrasSinAutorizar Then
                            'Reviso cual es la cantidad de hs que quedarian sin autorizar
                            StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                            StrSql = StrSql & " WHERE ternro = " & objRs!Ternro
                            StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha)
                            StrSql = StrSql & " AND thnro = " & no_autorizada
                            OpenRecordset StrSql, rsAD
                            If Not rsAD.EOF Then
                                If rsAD!adcanthoras >= MinimoExtrasSinAutorizar Then
                                    'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                                    StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                                    " WHERE ternro = " & objRs!Ternro & _
                                    " AND horfecrep = " & ConvFecha(Fecha) & _
                                    " AND thnro = " & no_autorizada
                                    objConn.Execute StrSql, , adExecuteNoRecords
                                Else
                                    'No genero la anormalidad
                                End If
                            Else
                                'el registro no está ==> no hago nada (esto no debiera suceder)
                            End If
                        Else
                            'Quiere decir que no está configurada la politica o no tiene alcance
                            'POr cuestiones de compatibilidad lo dejo como esaba en el estandar antes de esta modificacion
                            
                            'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                            StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                            " WHERE ternro = " & objRs!Ternro & _
                            " AND horfecrep = " & ConvFecha(Fecha) & _
                            " AND thnro = " & no_autorizada
                            objConn.Execute StrSql, , adExecuteNoRecords
                        End If
                        'FGZ - 07/05/2009 --------------------------------------------------------------
                    End If
                    rs.Close
                    
                Else
                    '/* El total de horas es autorizable */
                    horasaut = objRs!adcanthoras
                    
                    StrSql = "DELETE FROM gti_acumdiario WHERE" & _
                    " ternro = " & objRs!Ternro & _
                    " AND adfecha = " & ConvFecha(Fecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    
                    'FGZ 18/09/08 Debo actualizar la hora original por si se le habia generado la onormalidad anteriormente.
                    StrSql = "UPDATE gti_horcumplido SET normnro = 0 " & _
                    " WHERE ternro = " & objRs!Ternro & _
                    " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                    
                    
                    StrSql = "select * from gti_acumdiario where" & _
                    " ternro = " & NroTer & _
                    " and thnro = " & autorizada & _
                    " and adfecha = " & ConvFecha(Fecha)
                    OpenRecordset StrSql, rs
                    
                    If (horasaut > 0) Then
                        If (rs.EOF) Then
                            TotHorHHMM = CHoras(horasaut, 60)
                            
                            StrSql = "INSERT INTO gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
                            " values(" & NroTer & "," & autorizada & "," & _
                            ConvFecha(Fecha) & "," & TotHorHHMM & "," & horasaut & ")"
                        Else
                            TotHorHHMM = CHoras(rs!adcanthoras + horasaut, 60)
                            
                            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ", adcanthoras = " & _
                            rs!adcanthoras + horasaut & " where " & _
                            " ternro = " & objRs!Ternro & _
                            " and adfecha = " & ConvFecha(Fecha) & _
                            " and thnro = " & autorizada
                        End If
                        objConn.Execute StrSql, , adExecuteNoRecords
                    End If
                    rs.Close
                
                End If
                
                rsAutdet.MoveNext
            Loop
            
            If Not entro Then
                'cambiar en acumdiario el tipo de hora
                
                'objRs!thnro = no_autorizada
                StrSql = "update gti_acumdiario set thnro = " & _
                no_autorizada & " WHERE " & _
                " ternro = " & objRs!Ternro & _
                " and adfecha = " & ConvFecha(Fecha) & _
                " and thnro = " & objRs!thnro
                objConn.Execute StrSql, , adExecuteNoRecords
                
                
                'FGZ - 07/05/2009 --------------------------------------------------------------
                'Control sobre la cantidad minima de hs sin autorizar que deben quedar para que se genere la anormalidad
                If HayMinimoExtrasSinAutorizar Then
                    'Reviso cual es la cantidad de hs que quedarian sin autorizar
                    StrSql = "SELECT adcanthoras FROM gti_acumdiario "
                    StrSql = StrSql & " WHERE ternro = " & objRs!Ternro
                    StrSql = StrSql & " AND adfecha = " & ConvFecha(Fecha)
                    StrSql = StrSql & " AND thnro = " & objRs!thnro
                    OpenRecordset StrSql, rsAD
                    If Not rsAD.EOF Then
                        If rsAD!adcanthoras >= MinimoExtrasSinAutorizar Then
                            'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                            StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                            " WHERE ternro = " & objRs!Ternro & _
                            " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                            " AND thnro = " & objRs!thnro
                            objConn.Execute StrSql, , adExecuteNoRecords
                        Else
                            'No genero la anormalidad
                        End If
                    Else
                        'el registro no está ==> no hago nada (esto no debiera suceder)
                    End If
                Else
                    'Quiere decir que no está configurada la politica o no tiene alcance
                    'POr cuestiones de compatibilidad lo dejo como esaba en el estandar antes de esta modificacion
                    
                    'C.A.T 6/8/08 Si no existe el parte cargo la anormalidad de Extras no Autorizadas
                    StrSql = "UPDATE gti_horcumplido SET normnro = 11 " & _
                    " WHERE ternro = " & objRs!Ternro & _
                    " AND horfecrep = " & ConvFecha(objRs!adfecha) & _
                    " AND thnro = " & objRs!thnro
                    objConn.Execute StrSql, , adExecuteNoRecords
                End If
                'FGZ - 07/05/2009 --------------------------------------------------------------
            End If
        End If
    Else
        ' El tipo de Hora no tiene configurado Autoriza y no autoriza
        If depurar Then
            Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Hora NO tiene configurado Autoriza y No Autoriza: " & rs!thautpor
        End If
        'Exit Sub
    End If
NextProgress:
    objRs.MoveNext
Loop

If objRs.State = adStateOpen Then objRs.Close
If rsAutdet.State = adStateOpen Then rsAutdet.Close

'Se recorren todos los partes del tercero directos (crean directamente AD)
StrSql = "select * from gti_cabparte " & _
" INNER JOIN gti_autdet on gti_cabparte.gcpnro = gti_autdet.gcpnro " & _
"WHERE (gcpdesde <= " & ConvFecha(Fecha) & _
") and (gcphasta >= " & ConvFecha(Fecha) & ")" & _
" AND ternro = " & NroTer & _
" and gadautorizable = 0 and " & _
"((gadfecdesde <= " & ConvFecha(Fecha) & " or (gadfecdesde is null)) and " & _
"(gadfechasta >= " & ConvFecha(Fecha) & " or (gadfechasta is null)))"
OpenRecordset StrSql, rsAutdet
Do While Not rsAutdet.EOF

    OpenRecordset "select * from cystipo where cystipnro = 1", rs
    'Verificar si esta en el NIVEL FINAL DE FIRMA
    If rs!cystipact = -1 Then
        StrSql = "select * from cysfirmas where cysfirfin = -1 and " & _
        "cysfircodext = '" & rsAutdet!gcpnro & "' and cystipnro = 1"
        rs.Close
        OpenRecordset StrSql, rs
        If rs.EOF Then
            Firmado = False
        Else
            Firmado = True
        End If
    Else
        Firmado = True
    End If
    rs.Close
    
    If Firmado Then
        StrSql = "select * from gti_acumdiario where" & _
        " ternro = " & rsAutdet!Ternro & _
        " and thnro = " & rsAutdet!thnro & _
        " and adfecha = " & ConvFecha(Fecha)
        OpenRecordset StrSql, rs
        If rs.EOF Then
            TotHorHHMM = CHoras(rsAutdet!gadhoras, 60)
            StrSql = "INSERT INTO gti_acumdiario (ternro,thnro,adfecha,horas,adcanthoras)" & _
            " values(" & rsAutdet!Ternro & "," & rsAutdet!thnro & "," & _
            ConvFecha(Fecha) & "," & TotHorHHMM & "," & rsAutdet!gadhoras & ")"
            objConn.Execute StrSql, , adExecuteNoRecords
        Else
            TotHorHHMM = CHoras(rs!adcanthoras + rsAutdet!gadhoras, 60)
            StrSql = "UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ",adcanthoras = " & _
            rs!adcanthoras + rsAutdet!gadhoras & " where " & _
            " ternro = " & rsAutdet!Ternro & _
            " and adfecha = " & ConvFecha(Fecha) & _
            " and thnro = " & rsAutdet!thnro
            objConn.Execute StrSql, , adExecuteNoRecords
            Cant = Cant + 1
        End If
        rs.Close
    End If
    rsAutdet.MoveNext
Loop

End Sub




Public Function ContarHs(ByVal Tercero As Long, ByVal FechaDesde As Date, ByVal HoraDesde As String, ByVal FechaHasta As Date, ByVal HoraHasta As String) As Single
' ---------------------------------------------------------------------------------------------
' Descripcion: Funcion que cuenta la cantidad de horas que hay en el HC entre un rango de hs.
' Autor      : FGZ
' Fecha      : 15/02/2006
' Ultima Mod.:
' Descripcion:
' ---------------------------------------------------------------------------------------------
Dim rs_HC As New ADODB.Recordset
Dim CantidadHS As Single
Dim TotHor As Single
Dim Tdias As Single
Dim Thoras As Integer
Dim Tmin As Integer

            CantidadHS = 0
            
            StrSql = " SELECT * FROM gti_horcumplido "
            StrSql = StrSql & " WHERE ternro = " & Tercero
            StrSql = StrSql & " AND hordesde >= " & ConvFecha(FechaDesde)
            StrSql = StrSql & " AND horhasta <= " & ConvFecha(FechaHasta)
            StrSql = StrSql & " Order BY thnro, hornro, horhoradesde"
            If rs_HC.State = adStateOpen Then rs_HC.Close
            OpenRecordset StrSql, rs_HC
            
            Do While Not rs_HC.EOF
                'Registra   [---------------]
                'Horario          (----------------)
                If objFechasHoras.Menor_Hora(rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde, "2200") And objFechasHoras.Menor_Hora(FechaDesde, "2200", rs_HC!horhasta, rs_HC!horhorahasta) And objFechasHoras.Menor_Igual_Hora(rs_HC!horhasta, rs_HC!horhorahasta, FechaHasta, "0600") Then
                    objFechasHoras.RestaHs FechaDesde, "2200", rs_HC!horhasta, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
                    CantidadHS = CantidadHS + TotHor
                End If
    
                'Registra       [---------------]
                'Horario    (----------------)
                If objFechasHoras.Mayor_Igual_Hora(rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde, "2200") And objFechasHoras.Menor_Hora(rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde, hora_hasta) And objFechasHoras.Mayor_Hora(rs_HC!horhasta, rs_HC!horhorahasta, FechaHasta, "0600") Then
                    objFechasHoras.RestaHs rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde + 1, "0600", Tdias, Thoras, Tmin
                    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
                    CantidadHS = CantidadHS + TotHor
                End If
                
                'Registra       [---------------]
                'Horario    (-----------------------)
                If objFechasHoras.Mayor_Igual_Hora(rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde, "2200") And objFechasHoras.Menor_Hora(rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde, hora_hasta) And objFechasHoras.Menor_Igual_Hora(rs_HC!horhasta, rs_HC!horhorahasta, FechaDesde, hora_hasta) And objFechasHoras.Mayor_Hora(rs_HC!horhasta, rs_HC!horhorahasta, FechaDesde, "2200") Then
                    objFechasHoras.RestaHs rs_HC!hordesde, rs_HC!horhoradesde, rs_HC!horhasta, rs_HC!horhorahasta, Tdias, Thoras, Tmin
                    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
                    CantidadHS = CantidadHS + TotHor
                End If
    
                'Registra   [---------------]
                'Horario        (--------)
                If objFechasHoras.Menor_Hora(rs_HC!hordesde, rs_HC!horhoradesde, FechaDesde, "2200") And objFechasHoras.Mayor_Hora(rs_HC!horhasta, rs_HC!horhorahasta, FechaDesde + 1, "0600") Then
                'If (rs_HC!horhoradesde < "2200" And rs_HC!horhorahasta > hora_hasta) Then
                    objFechasHoras.RestaHs FechaDesde, "2200", FechaDesde + 1, "0600", Tdias, Thoras, Tmin
                    TotHor = (Tdias * 24) + (Thoras + (Tmin / 60))
                    
                    CantidadHS = CantidadHS + TotHor
                End If
        
                rs_HC.MoveNext
            Loop
            
            ContarHs = CantidadHS

End Function


Public Sub Restar_Horas(ByVal Fecha_Inicio As Date, ByVal hora_inicio As String, ByVal Fecha_Fin As Date, ByVal hora_fin As String, ByRef TotDias As Integer, ByRef tothoras As Integer, ByRef TotMin As Integer)
'------------------------------------------------------------------------------------------
'Descripcion:   Calcula la cantidad horas minutos y dias
'Autor:         FGZ
'Fecha:         04/07/2005
'Ult. Modif:
'------------------------------------------------------------------------------------------
Dim total As Integer
Dim cantdias  As Integer
Dim canthoras As Integer
Dim Dia   As Integer '  cantidad de minutos en un dia
Dim Hora As Integer   ' cantidad de minutos en una hora

    Dia = 1440
    Hora = 60
    canthoras = (Int(Mid(hora_fin, 1, 2)) * Hora + _
                   Int(Mid(hora_fin, 3, 2))) - _
                  (Int(Mid(hora_inicio, 1, 2)) * Hora + _
                   Int(Mid(hora_inicio, 3, 2)))
    cantdias = DateDiff("d", Fecha_Inicio, Fecha_Fin) * Dia
    
    total = cantdias + canthoras
    TotDias = Int(total / Dia)
    tothoras = Int((total Mod Dia) / Hora)
    TotMin = (total Mod Hora)
End Sub



Public Sub Prog_23_Turno_Plus(ByVal p_ternro As Long, ByVal p_fecha As Date, ByVal hd_thorigen As Long, ByVal hd_thdestino As Long, ByVal Cantidad As Single)
'  --------------------------------------------------------------------------------------------------
'  Archivo:
'  Descripción: Conversion CUSTOM TELEARTE
'  Autor : FGZ
'  --------------------------------------------------------------------------------------------------
'  Ult Modif:
'  --------------------------------------------------------------------------------------------------
Dim Horas_Oblig As Single
Dim HorasRes As Single
Dim TotHor As Single
Dim Hora_Dest As Single
Dim Hora_Ori As Single
Dim Cod_Convenio As String

Dim objRs As New ADODB.Recordset
Dim objRsAD As New ADODB.Recordset
Dim objrhest As New ADODB.Recordset

Dim Val_Comida As Single
Dim Val_Merienda As Single
Dim THVal_Comida As Long
Dim THVal_Merienda As Long

Dim THMedioTurno As Long
Dim THTurno As Long
Dim THTurnoyMedio As Long
Dim Val_MedioTurno As Single
Dim Val_Turno As Single
Dim Val_TurnoyMedio As Single

Dim TotHorHHMM As String


    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inicio Programa"
    End If
    On Error GoTo ME_Conv

     'Tipos de hora destino
     THMedioTurno = 57
     THTurno = 58
     THTurnoyMedio = 59
     
     StrSql = " SELECT estrcodext FROM his_estructura, estructura "
     StrSql = StrSql & " WHERE his_estructura.tenro = 19 and htethasta is null and ternro = " & p_ternro & " and "
     StrSql = StrSql & " his_estructura.estrnro = estructura.estrnro"
     OpenRecordset StrSql, objrhest
     If Not objrhest.EOF Then
         If Not EsNulo(objrhest!estrcodext) Then
             Cod_Convenio = objrhest!estrcodext
             If depurar Then
                 Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio del empleado: " & Cod_Convenio
             End If
         Else
             Cod_Convenio = ""
             If depurar Then
                 Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Convenio sin Codigo Externo. No aplica. "
             End If
         End If
     Else
         If depurar Then
             Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Empleado sin convenio. No aplica. "
         End If
     End If
    
     If UCase(Cod_Convenio) = "SAL" Then
         If depurar Then
             Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Aplica para el Convenio. "
         End If
         'revisar condiciones
         'a)De 0 hs a 3 horas excedentes generdas = 1 medio turno
         'b)De 3 hs a 6 horas excedentes generdas = 1 turno
         'c)De 6 hs a 10 horas excedentes generdas = 1 turno y medio
         StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & hd_thorigen & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
         OpenRecordset StrSql, objRsAD
         If objRsAD.EOF Then
             Hora_Ori = 0
         Else
             Hora_Ori = objRsAD!adcanthoras
         End If
         If depurar Then
             Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Horas excedentes " & Hora_Ori
         End If
         
         Val_MedioTurno = 0
         Val_Turno = 0
         Val_TurnoyMedio = 0
         
         Select Case Hora_Ori
         Case Is > 6
             Val_MedioTurno = 0
             Val_Turno = 0
             Val_TurnoyMedio = 1
         Case Is > 3
             Val_MedioTurno = 0
             Val_Turno = 1
             Val_TurnoyMedio = 0
         Case Is > 0
             Val_MedioTurno = 1
             Val_Turno = 0
             Val_TurnoyMedio = 0
         Case Else
             Val_MedioTurno = 0
             Val_Turno = 0
             Val_TurnoyMedio = 0
         End Select
         If depurar Then
             Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Corresponden. " & Val_MedioTurno & " Medio Turno, " & Val_Turno & " Turno y " & Val_TurnoyMedio & " Turno y Medio."
         End If
        
         'Medio Turno Plus
         If Val_MedioTurno > 0 Then
             StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & THMedioTurno & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
             OpenRecordset StrSql, objRsAD
             If Not objRsAD.EOF Then
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                 End If
                 
                 TotHorHHMM = CHoras(objRsAD!adcanthoras + Val_MedioTurno, 60)
                 
                 StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(Val_MedioTurno, 3)
                 StrSql = StrSql & " WHERE thnro = " & THMedioTurno & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                 End If
             Else
                 TotHorHHMM = CHoras(Val_MedioTurno, 60)
                 StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro, horas, adcanthoras,admanual,advalido) " & _
                          " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & THMedioTurno & "," & TotHorHHMM & "," & Round(Val_MedioTurno, 3) & "," & _
                          CInt(False) & "," & CInt(True) & ")"
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                 End If
             End If
         End If
         
         'Turno Plus
         If Val_Turno > 0 Then
             StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & THTurno & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
             OpenRecordset StrSql, objRsAD
             If Not objRsAD.EOF Then
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                 End If
                 TotHorHHMM = CHoras(objRsAD!adcanthoras + Val_Turno, 60)
                 StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(Val_Turno, 3)
                 StrSql = StrSql & " WHERE thnro = " & THTurno & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                 End If
             Else
                 TotHorHHMM = CHoras(Val_Turno, 60)
                 StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro, horas,adcanthoras,admanual,advalido) " & _
                          " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & THTurno & "," & TotHorHHMM & "," & Round(Val_Turno, 3) & "," & _
                          CInt(False) & "," & CInt(True) & ")"
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                 End If
             End If
         End If
         'Turno y Medio Plus
         If Val_TurnoyMedio > 0 Then
             StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & THTurnoyMedio & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
             OpenRecordset StrSql, objRsAD
             If Not objRsAD.EOF Then
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "  habia " & objRsAD!adcanthoras
                 End If
                 TotHorHHMM = CHoras(objRsAD!adcanthoras + Val_TurnoyMedio, 60)
                 StrSql = " UPDATE gti_acumdiario SET horas =" & TotHorHHMM & ",adcanthoras = adcanthoras + " & Round(Val_TurnoyMedio, 3)
                 StrSql = StrSql & " WHERE thnro = " & THTurnoyMedio & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Update " & StrSql
                 End If
             Else
                 TotHorHHMM = CHoras(Val_TurnoyMedio, 60)
                 StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro, horas,adcanthoras,admanual,advalido) " & _
                          " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & THTurnoyMedio & "," & TotHorHHMM & "," & Round(Val_TurnoyMedio, 3) & "," & _
                          CInt(False) & "," & CInt(True) & ")"
                 objConn.Execute StrSql, , adExecuteNoRecords
                 If depurar Then
                     Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Insert " & StrSql
                 End If
             End If
         End If
     Else
     End If

FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Fin Programa"
    End If
    
    'cierro todo
    If objrhest.State = adStateOpen Then objrhest.Close
    If objRs.State = adStateOpen Then objRs.Close
    If objRsAD.State = adStateOpen Then objRsAD.Close
    
    Set objrhest = Nothing
    Set objRs = Nothing
    Set objRsAD = Nothing
    Exit Sub
    
ME_Conv:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "**********************************************************"
    Flog.writeline
End Sub


'   1.  Conversion              : Conversion estandar a Jornada produccion
'   2.  ConversionProd          : Customizacion para Moño Azul
'   3.  SACO1HORA               : Customizacion para Temaiken
'   4.  REDONDEO                :
'   5.  SABADOS SCHERING        : Customizacion para Schering
'   6.  NormalesEstrada         : Customizacion para Estrada
'   7.  H100DIVINO              : Customizacion para Divino
'   8.  ConvNormales            : Customizacion para ICI
'   9.  Conv50%                 : Customizacion para ICI
'   10. Conv100%                : Customizacion para ICI
'   11. Conv200%                : Customizacion para ICI
'   12. Feriados                : Customizacion para AGD
'   13. Feriados_Estr           : Customizacion para AGD
'   14. Feriados_Trabajados     : Customizacion para AGD
'   15. Feriados_Trabajados_SD  : Customizacion para AGD
'   16. HorasDestajo            : Customizacion para Frig. Gorina
'   17. Adicalmuerzo            : Customizacion para Frig. Gorina
'   18. Peficiencia             : Customizacion para Frig. Gorina
'   19. Completar               : Customizacion para Schneider.
'   20. SABADODOMINGO MV        : Customizacion para MultiVoice.
'   21. TOPEMINIMO              : Customizacion para TRILENIUM extensible al estandar.
'   21. TOPEMINIMO_LV           : Customizacion para TRILENIUM extensible al estandar.
'   21. TOPEMINIMO_SD           : Customizacion para TRILENIUM extensible al estandar.
'   22. VALES_SAT               : Customizacion para TELEARTE.
'   23. TURNO_PLUS              : Customizacion para TELEARTE.



Public Sub Prog_1_Conversion(ByVal p_ternro As Long, ByVal p_fecha As Date, ByVal hd_thorigen As Long, ByVal hd_thdestino As Long, ByVal Cantidad As Single)
'  --------------------------------------------------------------------------------------------------
'  Archivo:
'  Descripción: Conversion CUSTOM
'  Autor : FGZ
'  --------------------------------------------------------------------------------------------------
'  Ult Modif:
'  --------------------------------------------------------------------------------------------------
Dim Horas_Oblig As Single
Dim HorasRes As Single
Dim TotHor As Single
Dim Hora_Dest As Single
Dim Hora_Ori As Single

Dim objRs As New ADODB.Recordset
Dim objRsAD As New ADODB.Recordset


Dim TotHorHHMM As String

If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inicio Programa"
End If
On Error GoTo ME_Conv


        'Programa que convierte de la Cantida de Horas en Cantidad de Días para el turno del Empleado
        StrSql = "SELECT * FROM gti_dias WHERE subturnro = " & Nro_Subturno
        StrSql = StrSql & " ORDER BY diaorden"
        OpenRecordset StrSql, objRs
        If Not objRs.EOF Then
            Horas_Oblig = objRs!diacanthoras
        End If
        
        If Horas_Oblig > 0 Then
            StrSql = " SELECT * FROM gti_acumdiario WHERE thnro = " & hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
            OpenRecordset StrSql, objRsAD
            If Not objRsAD.EOF Then
                If (objRsAD!adcanthoras / Horas_Oblig) < 1 Then
                    HorasRes = 1
                Else
                    HorasRes = objRsAD!adcanthoras / Horas_Oblig
                End If
            
                TotHorHHMM = CHoras(HorasRes, 60)
                StrSql = " UPDATE gti_acumdiario SET horas = " & TotHorHHMM & ", adcanthoras = " & Round(objRsAD!adcanthoras / Horas_Oblig, 3)
                StrSql = StrSql & " WHERE thnro = " & hd_thdestino & " AND ternro = " & p_ternro & " AND adfecha = " & ConvFecha(p_fecha)
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                TotHorHHMM = CHoras(Round(Cantidad / Horas_Oblig, 3), 60)
                StrSql = " INSERT INTO gti_acumdiario(adfecha,ternro,thnro, horas,adcanthoras,admanual,advalido) " & _
                         " VALUES (" & ConvFecha(p_fecha) & "," & p_ternro & "," & hd_thdestino & "," & TotHorHHMM & "," & Round(Cantidad / Horas_Oblig, 3) & "," & _
                         CInt(False) & "," & CInt(True) & ")"
                objConn.Execute StrSql, , adExecuteNoRecords
            End If
        Else
            If depurar Then
                Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Conversion abortada, Cantidad de horas produccion del turno es 0."
            End If
        End If


FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Fin Programa"
    End If
    
    'cierro todo
    If objRs.State = adStateOpen Then objRs.Close
    If objRsAD.State = adStateOpen Then objRsAD.Close
    
    Set objRs = Nothing
    Set objRsAD = Nothing
    Exit Sub
    
ME_Conv:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "**********************************************************"
    Flog.writeline
End Sub



Public Sub Prog_X_XX(ByVal p_ternro As Long, ByVal p_fecha As Date, ByVal hd_thorigen As Long, ByVal hd_thdestino As Long, ByVal Cantidad As Single)
'  --------------------------------------------------------------------------------------------------
'  Archivo:
'  Descripción: Conversion CUSTOM
'  Autor : FGZ
'  --------------------------------------------------------------------------------------------------
'  Ult Modif:
'  --------------------------------------------------------------------------------------------------
Dim Horas_Oblig As Single
Dim HorasRes As Single
Dim TotHor As Single
Dim Hora_Dest As Single
Dim Hora_Ori As Single
Dim Nro_Dire As Long
Dim Nro_Ccos As Long
Dim Nro_GSeg As Long
Dim RestoDecimal As Single

Dim EntroAntes11 As Boolean
Dim Total50 As Single
Dim Total100 As Single
Dim TotalHoras As Single

Dim Total_Antes13 As Single
Dim Total_Despues13 As Single

Dim TotalNocturnas As Single
Dim Total150 As Single

Dim TipoHora50 As Long
Dim TipoHora100 As Long
Dim TipoHora150 As Long
Dim TipoNocturna As Integer

Dim TipoHoraNoc100 As Long
Dim TipoHoraNoc150 As Long
Dim TipoHoraFer100 As Long
Dim TipoHoraFer150 As Long

Dim QuedanHs As Boolean
Dim SaldoHS As Single
Dim Dias As Integer
Dim Horas As Integer
Dim Minutos As Integer

Dim Limite1 As String
Dim Limite2 As String

Dim CCosto As Long
Dim Sector As Long
Dim Tenro As Long
Dim Cod_Convenio As String

Dim Tipos_de_Licencias As String
Dim Hay_Licencia As Boolean
Dim Rs_Justif As New ADODB.Recordset
Dim Rs_Lic As New ADODB.Recordset

Dim objRsCFG As New ADODB.Recordset
Dim objRsAD As New ADODB.Recordset
Dim objRsAD100 As New ADODB.Recordset
Dim objrhest As New ADODB.Recordset
Dim rs_HC As New ADODB.Recordset
Dim rs_AD As New ADODB.Recordset
Dim rs_Estructura As New ADODB.Recordset
Dim rs_Cab As New ADODB.Recordset
Dim SinConvenio As Boolean
Dim ConvenioAnterior As Boolean
Dim rs_TH As New ADODB.Recordset
Dim THOrigen As Long
Dim rs_ST As New ADODB.Recordset
Dim TH_Anormalidad As Long

Dim CantidadDestino As Single
Dim CantidadOrigen As Single
Dim Continua As Boolean

Dim Val_Comida As Single
Dim Val_Merienda As Single
Dim THVal_Comida As Long
Dim THVal_Merienda As Long

Dim THMedioTurno As Long
Dim THTurno As Long
Dim THTurnoyMedio As Long
Dim Val_MedioTurno As Single
Dim Val_Turno As Single
Dim Val_TurnoyMedio As Single

Dim TotHorHHMM As String

If depurar Then
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Inicio Programa"
End If
On Error GoTo ME_Conv











FIN:
    If depurar Then
        Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "Fin Programa"
    End If
    
    'cierro todo
    If objRsCFG.State = adStateOpen Then objRsCFG.Close
    If objRsAD.State = adStateOpen Then objRsAD.Close
    If objRsAD100.State = adStateOpen Then objRsAD100.Close
    If objrhest.State = adStateOpen Then objrhest.Close
    If rs_HC.State = adStateOpen Then rs_HC.Close
    If rs_AD.State = adStateOpen Then rs_AD.Close
    If rs_Estructura.State = adStateOpen Then rs_Estructura.Close
    If rs_Cab.State = adStateOpen Then rs_Cab.Close
    
    Set objRsCFG = Nothing
    Set objRsAD = Nothing
    Set objRsAD100 = Nothing
    Set objrhest = Nothing
    Set rs_HC = Nothing
    Set rs_AD = Nothing
    Set rs_Estructura = Nothing
    Set rs_Cab = Nothing
    Exit Sub
    
ME_Conv:
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "**********************************************************"
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & Err.Description
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "SQL Ejecutado: " & StrSql
    Flog.writeline
    Flog.writeline Espacios(Tabulador * Nivel_Tab_Log) & "**********************************************************"
    Flog.writeline
End Sub





