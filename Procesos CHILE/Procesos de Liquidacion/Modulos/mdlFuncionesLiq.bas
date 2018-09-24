Attribute VB_Name = "mdlFuncionesLiq"
'LED 26/08/2013 - Se agrego funcion buscarDescripcionAcuCon(tipo,codigo) busca la descripcion de un acumulador o concepto

'Busca un Concepto, Acumulador Mesual o Acumulador Liquidacion segun la etiqueta configurada en el confRep.
'(Concepto -> COC | COM) -- (Acumulador Mensual -> ACM | ACC) -- (Acumulador de Liquidacion -> ALC | ALM)
Function buscarConceptoAcumPorEtiqueta(ByVal busEtiq As String, ByVal Ternro As Long, ByVal codigoConAcu As String, ByVal mesLiq As Integer, ByVal anioLiq As Long)
 Dim rsValorLiq As New ADODB.Recordset
    
    Select Case busEtiq
        Case "COC", "COM" '-------------------------------------------------------------------------------------------------------------
            'Busco todos lod detliq entre los meses
            StrSql = "SELECT sum(detliq.dlicant) dlicant, sum(detliq.dlimonto) dlimonto  FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND cabliq.empleado = " & Ternro & _
                    " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
                    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.conccod='" & codigoConAcu & "'" & _
                    " WHERE periodo.pliqmes= " & mesLiq & "  AND periodo.pliqanio= " & anioLiq
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "COC" Then
                    buscarConceptoAcumPorEtiqueta = IIf(EsNulo(rsValorLiq!dlicant), 0, rsValorLiq!dlicant)
                Else
                    buscarConceptoAcumPorEtiqueta = IIf(EsNulo(rsValorLiq!dlimonto), 0, rsValorLiq!dlimonto)
                End If
            Else
                buscarConceptoAcumPorEtiqueta = 0
            End If
        
        Case "ACC", "ACM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT SUM(ammonto) ammonto, SUM(amcant) amcant FROM acu_mes WHERE acu_mes.ammes= " & mesLiq & "  AND acu_mes.amanio= " & anioLiq & " AND ternro = " & Ternro & _
                    " AND acunro= " & codigoConAcu
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ACC" Then
                    buscarConceptoAcumPorEtiqueta = IIf(EsNulo(rsValorLiq!amcant), 0, rsValorLiq!amcant)
                Else
                    buscarConceptoAcumPorEtiqueta = IIf(EsNulo(rsValorLiq!ammonto), 0, rsValorLiq!ammonto)
                End If
            Else
                buscarConceptoAcumPorEtiqueta = 0
            End If
        
        Case "ALC", "ALM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT sum(acu_liq.alcant) alcant, sum(acu_liq.almonto) almonto FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
                    " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                    " WHERE cabliq.empleado = " & Ternro & _
                    " AND periodo.pliqmes = " & mesLiq & _
                    " AND periodo.pliqanio = " & anioLiq & _
                    " AND acu_liq.acunro = " & codigoConAcu
            ' " AND periodo.pliqmes >= " & (mesLiq - 12)
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ALC" Then
                    buscarConceptoAcumPorEtiqueta = IIf(EsNulo(rsValorLiq!alcant), 0, rsValorLiq!alcant)
                Else
                    buscarConceptoAcumPorEtiqueta = IIf(EsNulo(rsValorLiq!almonto), 0, rsValorLiq!almonto)
                End If
            Else
                buscarConceptoAcumPorEtiqueta = 0
            End If
            
    End Select
        

End Function

'Busca un Concepto, Acumulador o Acumulador Liquidacion Anual segun la etiqueta configurada en el confRep, y por sexo (-1 masculino, 0 Femenino) para un empleado en cierta estructura.
'(Concepto -> COC | COM) -- (Acumulador Mensual -> ACM | ACC) -- (Acumulador de Liquidacion -> ALC | ALM)
Function buscarConceptoAcumPorEtiquetaAnual(ByVal busEtiq As String, ByVal codigoConAcu As String, ByVal anioLiq As Long, ByVal sexo As Integer, ByVal estrnro As Long, ByVal fechaDesde As Date, ByVal fechaHasta As Date, ByVal empresa As Long, ByVal sucursal As Long)
 Dim rsValorLiq As New ADODB.Recordset
    
    Select Case busEtiq
        Case "COC", "COM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT sum(detliq.dlicant) dlicant, sum(detliq.dlimonto) dlimonto FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro  " & _
                    " INNER JOIN tercero ON cabliq.empleado = tercero.ternro and tersex = " & sexo & _
                    " INNER JOIN his_estructura Emp ON Emp.ternro = tercero.ternro and Emp.estrnro = " & empresa & _
                    " AND ((Emp.htetdesde <= " & ConvFecha(fechaDesde) & " AND (Emp.htethasta IS NULL OR  Emp.htethasta>= " & ConvFecha(fechaDesde) & "))" & _
                    " OR ((Emp.htetdesde >= " & ConvFecha(fechaDesde) & " AND Emp.htetdesde<= " & ConvFecha(fechaHasta) & ") ) ) " & _
                    " INNER JOIN his_estructura Suc ON Suc.ternro = tercero.ternro and Suc.estrnro = " & sucursal & _
                    " AND ((Suc.htetdesde <= " & ConvFecha(fechaDesde) & " AND (Suc.htethasta IS NULL OR  Suc.htethasta>= " & ConvFecha(fechaDesde) & "))" & _
                    " OR ((Suc.htetdesde >= " & ConvFecha(fechaDesde) & " AND Suc.htetdesde<= " & ConvFecha(fechaHasta) & "))) " & _
                    " INNER JOIN his_estructura est ON tercero.ternro = est.ternro and est.estrnro = " & estrnro & _
                    " AND ((est.htetdesde <= " & ConvFecha(fechaDesde) & "  AND (est.htethasta IS NULL OR  est.htethasta>= " & ConvFecha(fechaDesde) & " )) " & _
                    " OR ((est.htetdesde >= " & ConvFecha(fechaDesde) & "  AND est.htetdesde<= " & ConvFecha(fechaHasta) & " ) ) ) " & _
                    " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
                    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.conccod='" & codigoConAcu & "'" & _
                    " WHERE periodo.pliqanio= " & anioLiq
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "COC" Then
                    buscarConceptoAcumPorEtiquetaAnual = IIf(EsNulo(rsValorLiq!dlicant), 0, rsValorLiq!dlicant)
                Else
                    buscarConceptoAcumPorEtiquetaAnual = IIf(EsNulo(rsValorLiq!dlimonto), 0, rsValorLiq!dlicant)
                End If
            Else
                buscarConceptoAcumPorEtiquetaAnual = 0
            End If
        
        Case "ACC", "ACM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT sum(amcant) amcant, sum(ammonto) ammonto FROM acu_mes " & _
                    " INNER JOIN tercero ON acu_mes.ternro = tercero.ternro and tersex =  " & sexo & _
                    " INNER JOIN his_estructura Emp ON Emp.ternro = tercero.ternro and Emp.estrnro = " & empresa & _
                    " AND ((Emp.htetdesde <= " & ConvFecha(fechaDesde) & " AND (Emp.htethasta IS NULL OR  Emp.htethasta>= " & ConvFecha(fechaDesde) & "))" & _
                    " OR ((Emp.htetdesde >= " & ConvFecha(fechaDesde) & " AND Emp.htetdesde<= " & ConvFecha(fechaHasta) & ") ) ) " & _
                    " INNER JOIN his_estructura Suc ON Suc.ternro = tercero.ternro and Suc.estrnro = " & sucursal & _
                    " AND ((Suc.htetdesde <= " & ConvFecha(fechaDesde) & " AND (Suc.htethasta IS NULL OR  Suc.htethasta>= " & ConvFecha(fechaDesde) & "))" & _
                    " OR ((Suc.htetdesde >= " & ConvFecha(fechaDesde) & " AND Suc.htetdesde<= " & ConvFecha(fechaHasta) & "))) " & _
                    " INNER JOIN his_estructura est ON tercero.ternro = est.ternro and est.estrnro = " & estrnro & _
                    " AND ((est.htetdesde <= " & ConvFecha(fechaDesde) & "  AND (est.htethasta IS NULL OR  est.htethasta>= " & ConvFecha(fechaDesde) & " )) " & _
                    " OR ((est.htetdesde >= " & ConvFecha(fechaDesde) & "  AND est.htetdesde<= " & ConvFecha(fechaHasta) & " ) ) ) " & _
                    " WHERE acu_mes.amanio= " & anioLiq & " AND acunro= " & codigoConAcu
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ACC" Then
                    buscarConceptoAcumPorEtiquetaAnual = IIf(EsNulo(rsValorLiq!amcant), 0, rsValorLiq!amcant)
                Else
                    buscarConceptoAcumPorEtiquetaAnual = IIf(EsNulo(rsValorLiq!ammonto), 0, rsValorLiq!ammonto)
                End If
            Else
                buscarConceptoAcumPorEtiquetaAnual = 0
            End If
        
        Case "ALC", "ALM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT sum(acu_liq.alcant) alcant, sum(acu_liq.almonto) almonto FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro " & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
                    " INNER JOIN tercero ON cabliq.empleado = tercero.ternro and tersex =  " & sexo & _
                    " INNER JOIN his_estructura Emp ON Emp.ternro = tercero.ternro and Emp.estrnro = " & empresa & _
                    " AND ((Emp.htetdesde <= " & ConvFecha(fechaDesde) & " AND (Emp.htethasta IS NULL OR  Emp.htethasta>= " & ConvFecha(fechaDesde) & "))" & _
                    " OR ((Emp.htetdesde >= " & ConvFecha(fechaDesde) & " AND Emp.htetdesde<= " & ConvFecha(fechaHasta) & ") ) ) " & _
                    " INNER JOIN his_estructura Suc ON Suc.ternro = tercero.ternro and Suc.estrnro = " & sucursal & _
                    " AND ((Suc.htetdesde <= " & ConvFecha(fechaDesde) & " AND (Suc.htethasta IS NULL OR  Suc.htethasta>= " & ConvFecha(fechaDesde) & "))" & _
                    " OR ((Suc.htetdesde >= " & ConvFecha(fechaDesde) & " AND Suc.htetdesde<= " & ConvFecha(fechaHasta) & "))) " & _
                    " INNER JOIN his_estructura est ON tercero.ternro = est.ternro and est.estrnro = " & estrnro & _
                    " AND ((est.htetdesde <= " & ConvFecha(fechaDesde) & "  AND (est.htethasta IS NULL OR  est.htethasta>= " & ConvFecha(fechaDesde) & " )) " & _
                    " OR ((est.htetdesde >= " & ConvFecha(fechaDesde) & "  AND est.htetdesde<= " & ConvFecha(fechaHasta) & " ) ) ) " & _
                    " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                    " WHERE periodo.pliqanio = " & anioLiq & _
                    " AND acu_liq.acunro = " & codigoConAcu
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ALC" Then
                    buscarConceptoAcumPorEtiquetaAnual = IIf(EsNulo(rsValorLiq!alcant), 0, rsValorLiq!alcant)
                Else
                    buscarConceptoAcumPorEtiquetaAnual = IIf(EsNulo(rsValorLiq!almonto), 0, rsValorLiq!almonto)
                End If
            Else
                buscarConceptoAcumPorEtiquetaAnual = 0
            End If
            
    End Select
        

End Function

Function buscarDescripcionAcuCon(ByVal tipo As String, ByVal codigo As String)
'LED 26/08/2013 - Se agrego funcion buscarDescripcionAcuCon(tipo,codigo) busca la descripcion de un acumulador o concepto.
Dim rsDesc As New ADODB.Recordset
    
    Select Case tipo
        Case "COC", "COM"
            StrSql = "SELECT concabr FROM concepto WHERE conccod = '" & codigo & "'"
            OpenRecordset StrSql, rsDesc
            If Not rsDesc.EOF Then
                buscarDescripcionAcuCon = IIf(EsNulo(rsDesc!concabr), "", rsDesc!concabr)
            Else
                buscarDescripcionAcuCon = ""
            End If

        Case "ACC", "ACM", "ALC", "ALM"
            StrSql = "SELECT acudesabr FROM acumulador WHERE acunro = " & codigo
            OpenRecordset StrSql, rsDesc
            If Not rsDesc.EOF Then
                buscarDescripcionAcuCon = IIf(EsNulo(rsDesc!acudesabr), "", rsDesc!acudesabr)
            Else
                buscarDescripcionAcuCon = ""
            End If
        End Select
End Function

'Busca un Concepto, Acumulador Mesual o Acumulador Liquidacion segun la etiqueta configurada en el confRep.
'(Concepto -> COC | COM) -- (Acumulador Mensual -> ACM | ACC) -- (Acumulador de Liquidacion -> ALC | ALM)
Function buscarConceptoAcumPorEtiquetaEnProcesos(ByVal busEtiq As String, ByVal Ternro As Long, ByVal codigoConAcu As String, ByVal mesLiq As Integer, ByVal anioLiq As Long, ByVal procesos As String)
 Dim rsValorLiq As New ADODB.Recordset
    
    Select Case busEtiq
        Case "COC", "COM" '-------------------------------------------------------------------------------------------------------------
            'Busco todos lod detliq entre los meses
            StrSql = "SELECT sum(detliq.dlicant) dlicant, sum(detliq.dlimonto) dlimonto  FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro AND pronro in (" & procesos & ")" & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro AND cabliq.empleado = " & Ternro & _
                    " INNER JOIN detliq ON cabliq.cliqnro = detliq.cliqnro " & _
                    " INNER JOIN concepto ON concepto.concnro = detliq.concnro AND concepto.conccod='" & codigoConAcu & "'" & _
                    " WHERE periodo.pliqmes= " & mesLiq & "  AND periodo.pliqanio= " & anioLiq
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "COC" Then
                    buscarConceptoAcumPorEtiquetaEnProcesos = IIf(EsNulo(rsValorLiq!dlicant), 0, rsValorLiq!dlicant)
                Else
                    buscarConceptoAcumPorEtiquetaEnProcesos = IIf(EsNulo(rsValorLiq!dlimonto), 0, rsValorLiq!dlimonto)
                End If
            Else
                buscarConceptoAcumPorEtiquetaEnProcesos = 0
            End If
        
        Case "ACC", "ACM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT SUM(ammonto) ammonto, SUM(amcant) amcant FROM acu_mes WHERE acu_mes.ammes= " & mesLiq & "  AND acu_mes.amanio= " & anioLiq & " AND ternro = " & Ternro & _
                    " AND acunro= " & codigoConAcu
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ACC" Then
                    buscarConceptoAcumPorEtiquetaEnProcesos = IIf(EsNulo(rsValorLiq!amcant), 0, rsValorLiq!amcant)
                Else
                    buscarConceptoAcumPorEtiquetaEnProcesos = IIf(EsNulo(rsValorLiq!ammonto), 0, rsValorLiq!ammonto)
                End If
            Else
                buscarConceptoAcumPorEtiquetaEnProcesos = 0
            End If
        
        Case "ALC", "ALM" '-------------------------------------------------------------------------------------------------------------
            StrSql = "SELECT sum(acu_liq.alcant) alcant, sum(acu_liq.almonto) almonto FROM periodo " & _
                    " INNER JOIN proceso ON periodo.pliqnro = proceso.pliqnro AND pronro in (" & procesos & ")" & _
                    " INNER JOIN cabliq ON proceso.pronro = cabliq.pronro " & _
                    " INNER JOIN acu_liq ON cabliq.cliqnro = acu_liq.cliqnro " & _
                    " WHERE cabliq.empleado = " & Ternro & _
                    " AND periodo.pliqmes = " & mesLiq & _
                    " AND periodo.pliqanio = " & anioLiq & _
                    " AND acu_liq.acunro = " & codigoConAcu
            ' " AND periodo.pliqmes >= " & (mesLiq - 12)
            OpenRecordset StrSql, rsValorLiq
            
            If Not rsValorLiq.EOF Then
                If busEtiq = "ALC" Then
                    buscarConceptoAcumPorEtiquetaEnProcesos = IIf(EsNulo(rsValorLiq!alcant), 0, rsValorLiq!alcant)
                Else
                    buscarConceptoAcumPorEtiquetaEnProcesos = IIf(EsNulo(rsValorLiq!almonto), 0, rsValorLiq!almonto)
                End If
            Else
                buscarConceptoAcumPorEtiquetaEnProcesos = 0
            End If
            
    End Select
        

End Function

