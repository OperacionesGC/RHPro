create table histliq 
  (
    periodo_pago char(7) not null ,
    nro_corr_liq smallint not null ,
    fec_liq datetime,
    hora_liq char(8) not null ,
    descripcion varchar(255),
    cant_a_liquidar smallint,
    liq_realizada char(1) not null ,
    cant_liquidada smallint,
    liq_confirmada char(1) not null ,
    fec_confirmacion datetime,
    fec_pago datetime,
    fec_deposito datetime,
    fec_acredito datetime,
    clase_proc_liq char(1),
    proceso_de_ajuste char(1),
    periodo_pago_orig char(6),
    nro_corr_liq_orig smallint,
    tiene_resguardo char(1),
    grupo char(4),
    lista_empresa varchar(255),
    liq_en_proceso char(1)
  );

create unique index  i_hili on  histliq (periodo_pago,
    nro_corr_liq);


create table  legaliq 
  (
    periodo_pago char(8) not null ,
    nro_corr_liq smallint not null ,
    nro_leg integer not null ,
    tipo_liq char(4) not null ,
    liquidado char(1) not null ,
    confirmado char(1) not null ,
    estr_liq char(10),
    empresa char(4),
    ajuste_realizado char(1),
    fec_inicio_ajuste datetime,
    ult_dia_trabajado datetime,
    modo_extincion char(4),
    indem_preaviso char(1),
    dias_vaca_no_gozad smallint,
    compens_maternidad char(1),
    indem_esp_matrim char(1),
    indem_esp_maternid char(1),
    indic_indem_antig char(1),
    fecha_desde_vaca datetime,
    cant_dias_vacacion smallint,
    activa_cptos_futur char(1),
    mejor_remu_sem_sac decimal(15,5),
    ult_sueldo_salario decimal(15,5),
    prom_remun_variab decimal(15,5),
    mejor_remu_p_antig decimal(15,5)
  );

create unique index  i_lega on  legaliq (periodo_pago,
    nro_corr_liq,nro_leg,tipo_liq);


create table  histcon 
  (
    empresa char(4) not null ,
    periodo_pago char(7) not null ,
    nro_corr_liq smallint not null ,
    nro_leg integer not null ,
    cod_cpto char(8),
    nro_corr_cpto smallint not null ,
    tipo_liq char(4) not null ,
    estr_liq char(8) not null ,
    fec_liq datetime not null ,
    fec_concepto datetime,
    valor_base decimal(15,5),
    cantidad decimal(15,5),
    porcentaje decimal(6,3),
    importe_final decimal(15,3),
    valor_base_deveng decimal(15,5),
    cantidad_deveng decimal(15,5),
    importe_devengado money(15,3),
    centro_costo char(24),
    mot_dif_subsidio char(4),
    fec_ultima_act datetime,
    periodo_pago_deven char(6),
    nro_corr_liq_deven smallint,
    cod_cpto_deven char(8)
  );

create unique index  i_hicp on  histcon (periodo_pago,
    nro_corr_liq,nro_leg,cod_cpto,nro_corr_cpto,tipo_liq);
create index  i_recacum on  histcon (periodo_pago,
    nro_leg);


create table  concepto_defin 
  ( estr_liq char(8) not null , 
    cod_cpto char(8) not null ,
    nombre char(30),
    nombre_usuario char(30),
    descripcion varchar(255),
    texto_recibo char(30),
    id_algo char(50),
    orden_calc smallint,
    ing_valor_base char(1),
    ing_cantidad char(1),
    ing_porcentaje char(1),
    asume_cp_sin_vto char(1),
    acepta_duplicados char(1),
    junta_duplicados char(1),
    muestra_valor_base char(1),
    muestra_cantidad char(1),
    muestra_porcentaje char(1),
    min_valor_base decimal(15,5),
    max_valor_base decimal(15,5),
    min_cantidad decimal(15,5),
    max_cantidad decimal(15,5),
    min_porcentaje decimal(6,3),
    max_porcentaje decimal(6,3),
    min_importe_final decimal(15,3),
    max_importe_final decimal(15,3),
    redond_imp_final char(4),
    grupo char(4),
    fec_creacion datetime,
    hora_creacion char(8),
    fec_ult_actualiz datetime,
    hora_ult_actualiz char(8),
    id_usuario char(16),
    estr_liq_global char(1),
    estr_liq_origen char(1),
    tipo_obse char(4),
    importancia_obse smallint,
    observacion byte);

create unique index  i_cpto_def on  concepto_defin 
    (estr_liq,cod_cpto);
create index  i_cpto2 on  concepto_defin (muestra_valor_base,
    id_algo);
create index  i_cpto3 on  concepto_defin (muestra_cantidad,
    id_algo);
create index  i_cpto4 on  concepto_defin (muestra_porcentaje,
    id_algo);
create index  i_cpto5 on  concepto_defin (estr_liq);
    




