/*-----------------------------------------------------------------------------

  Archivo: suesin04.p
  Descripci¢n: llena la tabla para el reporte de sindicatos UOCRA y UOMRA.
  Autor: Marchese, Juan M.
  Creado: 5/5/2000
  Modificado:
  
-----------------------------------------------------------------------------*/


  /* Par metros */

  def input parameter nro-sindicato as int.
  def input parameter nroliq as int.
  def input parameter nropro as int.
  def input parameter todos-pro as log.
  def input parameter proc-aprob as log.
  def input parameter agrupado as char.


  /* Variables */

  def var sueldo as dec init 0.
  def var aporte as dec init 0.
  def var contribucion as dec init 0.


  {def-rep.i}  /* define la variable nro-reporte como compartida */     
  {headcom.i}


  /* Programa */

  find first per.periodo where per.periodo.pliqnro = nroliq no-lock no-error.
  if not avail(per.periodo)
  then
    leave.

  case agrupado :
  
    when "UOCRA" then
      do:
        find first per.reporte where per.reporte.repnro = 45 no-lock no-error.
        if not avail(per.reporte)
        then
          do:
            message "Error en Configuraci¢n de Reporte UOCRA." view-as alert-box error title "Error".
            leave.
          end.
      end.

    when "UOMRA" then
      do:
        find first per.reporte where per.reporte.repnro = 46 no-lock no-error.
        if not avail(per.reporte)
        then
          do:
            message "Error en Configuraci¢n de Reporte UOMRA." view-as alert-box error title "Error".
            leave.
          end.
      end.

    otherwise
      do:
        message "Error de Par metros." view-as alert-box title "Error".
        leave.
      end.

  end. /* case grupado */
  

  for each per.proceso of per.periodo where ((per.proceso.pronro = nropro or todos-pro or
                                             (per.proceso.proaprob and proc-aprob))) no-lock,
      each per.cabliq of per.proceso no-lock,
      each per.empleado where per.empleado.ternro = per.cabliq.empleado no-lock,
      each per.gremio where (per.gremio.ternro = per.empleado.gremio) and
                            ((per.gremio.ternro = nro-sindicato) or
                             (nro-sindicato = ?)) no-lock
/*                                  
  for each per.gremio where (per.gremio.ternro = nro-sindicato) or
                            (nro-sindicato = ?) no-lock,
      each per.empleado where per.empleado.gremio = per.gremio.ternro no-lock,
      each per.cabliq where per.cabliq.empleado = per.empleado.ternro no-lock,
      each per.proceso of per.periodo where (per.proceso.pronro = per.cabliq.pronro) and
                                            ((per.proceso.pronro = nropro or todos-pro or
                                             (per.proceso.proaprob and proc-aprob))) no-lock
*/
      break by per.empleado.ternro :


      for each per.confrep where per.confrep.repnro = per.reporte.repnro no-lock :
    
        if per.confrep.conftipo = 'AC'
        then
          do:
            find first per.acu_liq of per.cabliq where per.acu_liq.acunro = per.confrep.confval no-lock no-error.
            if available(per.acu_liq)
            then
              assign sueldo = sueldo + per.acu_liq.almonto.
          end.

        if confrep.conftipo = 'AP'
        then
          do:                                      
            find first per.concepto where per.concepto.conccod = string(per.confrep.confval, "9999") no-lock no-error.
            if available(per.concepto)
            then
              do:
                find first per.detliq of per.cabliq where per.detliq.concnro = per.concepto.concnro no-lock no-error.
                if available(per.detliq)
                then
                  assign aporte = aporte + per.detliq.dlimonto.

                find first per.prevliq of per.cabliq where per.prevliq.concnro = per.concepto.concnro no-lock no-error.
                if available(per.prevliq)
                then
                  assign aporte = aporte + per.prevliq.dlimonto.

                find first per.conliq of per.cabliq where per.conliq.concnro = per.concepto.concnro no-lock no-error.
                if available(per.conliq)
                then
                  assign aporte = aporte + per.conliq.dlimonto.
              end.    
          end.

        if confrep.conftipo = 'CO'
        then
          do:                                      
            find first per.concepto where per.concepto.conccod = string(per.confrep.confval, "9999") no-lock no-error.
            if available(per.concepto)
            then
              do:
                find first per.detliq of per.cabliq where per.detliq.concnro = per.concepto.concnro no-lock no-error.
                if available(per.detliq)
                then
                  assign contribucion = contribucion + per.detliq.dlimonto.

                find first per.prevliq of per.cabliq where per.prevliq.concnro = per.concepto.concnro no-lock no-error.
                if available(per.prevliq)
                then
                  assign contribucion = contribucion + per.prevliq.dlimonto.

                find first per.conliq of per.cabliq where per.conliq.concnro = per.concepto.concnro no-lock no-error.
                if available(per.conliq)
                then
                  assign contribucion = contribucion + per.conliq.dlimonto.
              end.    
          end.

      end. /* for each per.confrep */

      if last-of(per.empleado.ternro)
      then
        do:
        
          {busreg.i &Nro-Rep = 06}      

          assign rep06.empleado = aporte
                 rep06.empleador = contribucion
                 rep06.bruto = sueldo
                 rep06.sindicato = per.gremio.ternro
                 rep06.ternro = per.empleado.ternro.
          assign sueldo = 0
                 aporte = 0
                 contribucion = 0.
        end.

  end. /* for each per.gremio */

