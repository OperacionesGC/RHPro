Attribute VB_Name = "mdlPed_Argentina"
Public Sub GeneraPedido_ARG(ByVal fecha_desde, ByVal vacnro, ByVal vacdesc As String, ByVal alcannivel As Integer, ByVal Reproceso As Boolean)
    Dim rs_vacdiascor As New ADODB.Recordset
    Dim rsVac As New ADODB.Recordset
    Dim rsDias As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    diascoract = 0
    DiasTom = 0
    diascorant = 0
    diasdebe = 0
    diastot = 0
    diasyaped = 0
    diaspend = 0
    
    Flog.Writeline "Periodo de Vacaciones:" & vacnro & " " & vacdesc
    
    NroVac = vacnro
    
    'EAM- Obtiene los días correspondientes
    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
    StrSql = StrSql & " AND (venc = 0 OR venc IS NULL)"
    OpenRecordset StrSql, rs_vacdiascor
    If Not rs_vacdiascor.EOF Then
        diascoract = rs_vacdiascor!vdiascorcant ' dias corresp al periodo actual
        nroTipvac = rs_vacdiascor!tipvacnro
    Else
        diascoract = 0
    End If
    
    'Resto los vencidos
    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
    StrSql = StrSql & " AND (venc = 1)"
    OpenRecordset StrSql, rs_vacdiascor
    If Not rs_vacdiascor.EOF Then
        diascoract = diascoract - rs_vacdiascor!vdiascorcant
    End If
    
    'Sumo los transferidos
    StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
    StrSql = StrSql & " AND (venc = 2)"
    OpenRecordset StrSql, rs_vacdiascor
    'EAM- Dias tranferidos al periodo actual
    dias_tranf_PAct = 0
    DiasTom = 0
    
    If Not rs_vacdiascor.EOF Then
        diascoract = diascoract + rs_vacdiascor!vdiascorcant
        dias_tranf_PAct = rs_vacdiascor!vdiascorcant
    End If
    
    
    If diascoract > 0 Then
        'StrSql = "SELECT * FROM vacacion WHERE vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(fecha_desde)
        'EAM- Obtiene todos los periodos abiertos para el empleado en orden desc.
        StrSql = "SELECT DISTINCT vacacion.vacnro, vacdesc, vacfecdesde, vacfechasta,vacanio " & _
                " FROM vacacion " & _
                " INNER JOIN vac_estr ON vacacion.vacnro= vac_estr.vacnro" & _
                " INNER JOIN vacdiascor ON vac_estr.vacnro = vacdiascor.vacnro" & _
                " INNER JOIN his_estructura ON vac_estr.estrnro = his_estructura.estrnro " & _
                " WHERE his_estructura.ternro= " & Ternro & " AND vacacion.vacestado= -1 AND " & _
                " vacacion.vacnro <> " & NroVac & " and vacfechasta < " & ConvFecha(fecha_desde) & _
                " AND (venc = 1) " & _
                " ORDER BY vacanio DESC "
                 OpenRecordset StrSql, rsVac
        Do While Not rsVac.EOF
            DiasTom = 0
             StrSql = "SELECT * FROM lic_vacacion " & _
                      " INNER JOIN emp_lic ON emp_lic.emp_licnro = lic_vacacion.emp_licnro " & _
                      " WHERE lic_vacacion.vacnro = " & rsVac!vacnro & " AND emp_lic.empleado = " & Ternro
             OpenRecordset StrSql, rsDias
             Do While Not rsDias.EOF
                 DiasTom = DiasTom + rsDias!elcantdias
                 rsDias.MoveNext
             Loop
             
             'Busco los correspondientes al periodo
             If dias_tranf_PAct <> 0 Then
                 diascorant = (dias_tranf_PAct * (-1))
                 dias_tranf_PAct = 0
             Else
                 diascorant = 0
             End If
             
             'EAM- Obtine los días correspondientes del periodo
             StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
             StrSql = StrSql & " AND (venc = 0 OR venc IS NULL)"
             OpenRecordset StrSql, rs
             If Not rs.EOF Then
                 diascorant = diascorant + rs!vdiascorcant
                 
                 'resto los vencidos
                 StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
                 StrSql = StrSql & " AND (venc = 1)"
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     diascorant = diascorant - rs!vdiascorcant
                 End If
                 
                 'sumo los transferidos
                 StrSql = "SELECT * FROM vacdiascor WHERE ternro = " & Ternro & " AND vacnro = " & rsVac!vacnro
                 StrSql = StrSql & " AND (venc = 2)"
                 OpenRecordset StrSql, rs
                 If Not rs.EOF Then
                     'diascorant = diascorant + rs!vdiascorcant
                     dias_tranf_PAct = rs!vdiascorcant
                 End If
             Else
                 diascorant = 0
             End If
             
             
             diasdebe = diasdebe + (diascorant - DiasTom)
             
             rsVac.MoveNext
         Loop
         diastot = diascoract + diasdebe
    End If
    
    
    If Not Reproceso Then
        'Busco los pedidos de ese periodo
        StrSql = "SELECT * FROM vacdiasped WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
        OpenRecordset StrSql, objRs
        Do While Not objRs.EOF
            'diasyaped = diasyaped + objRs!vdiapedcant
            diasyaped = diasyaped + objRs!vdiaspedhabiles
            Aux_Fecha_Desde = IIf(Aux_Fecha_Desde < (objRs!vdiapedhasta + 1), objRs!vdiapedhasta + 1, Aux_Fecha_Desde)
            objRs.MoveNext
        Loop
    Else
        'borro los que estan en el rango de fechas
        StrSql = "DELETE FROM vacdiasped WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
        StrSql = StrSql & " AND vdiapeddesde >= " & ConvFecha(fecha_desde)
        objConn.Execute StrSql, , adExecuteNoRecords
        Flog.Writeline "Se Borraron por reprocesamiento los días pedidos del período " & NroVac & " >= a la fecha " & fecha_desde
        
        ' Busco los pedidos de ese periodo que quedaron afuera del rango de fechas
        StrSql = "SELECT * FROM vacdiasped WHERE ternro = " & Ternro & " AND vacnro = " & NroVac
        OpenRecordset StrSql, objRs
        Do While Not objRs.EOF
            'diasyaped = diasyaped + objRs!vdiapedcant
            diasyaped = diasyaped + objRs!vdiaspedhabiles
            'Aux_Fecha_Desde = objRs!vdiapedhasta + 1
            objRs.MoveNext
        Loop
        
    End If
    
    diaspend = diastot - diasyaped
    If diaspend > 0 Then
        If Todos_Posibles Then
            Call DiasPedidos_STD(nroTipvac, Aux_Fecha_Desde, hasta, Ternro, diaspend, tothabiles, totNohabiles, totferiados)
             
            'Verificar Fase
            If activo(Ternro, Aux_Fecha_Desde, hasta) Then
                StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
                          ConvFecha(hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Ternro & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
                objConn.Execute StrSql, , adExecuteNoRecords
            Else
                Flog.Writeline "No se insertaron los días " & Aux_Fecha_Desde & " a " & hasta & " porque se superpone con un período inactivo del empleado."
             End If
            
            Aux_Fecha_Desde = hasta + 1
        Else
             If Aux_Cant_dias > 0 Then
                If diaspend >= Aux_Cant_dias Then
                    diaspend = Aux_Cant_dias
                    Aux_Cant_dias = 0
                Else
                    Aux_Cant_dias = Aux_Cant_dias - diaspend
                End If
                'Call DiasPedidos(nroTipvac, fecha_desde, hasta, Ternro, diaspend, tothabiles, totNohabiles, totferiados)
                 Call DiasPedidos_STD(nroTipvac, Aux_Fecha_Desde, hasta, Ternro, diaspend, tothabiles, totNohabiles, totferiados)
           
                 'Verificar Fase
                If activo(Ternro, Aux_Fecha_Desde, hasta) Then
                     StrSql = "INSERT INTO vacdiasped (vdiapedhasta,vdiaspedferiados,vdiaspedhabiles,vdiaspednohabiles,ternro,vacnro,vdiapedcant,vdiapeddesde,vdiaspedestado) VALUES (" & _
                              ConvFecha(hasta) & "," & totferiados & "," & tothabiles & "," & totNohabiles & "," & Ternro & "," & NroVac & "," & (totferiados + tothabiles + totNohabiles) & "," & ConvFecha(Aux_Fecha_Desde) & ",-1)"
                    objConn.Execute StrSql, , adExecuteNoRecords
                Else
                    Flog.Writeline "No se insertaron los días " & Aux_Fecha_Desde & " a " & hasta & " porque se superpone con un período inactivo del empleado."
                End If
                
                Aux_Fecha_Desde = hasta + 1
            End If
        End If
    End If

End Sub
