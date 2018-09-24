CREATE TABLE anrtipofact  
( 
  tipfacnro integer IDENTITY (1, 1) not null,
  tipfacdesabr varchar(30),
  tipfacdesext varchar(200),
  tipfacestado smallint default -1 not null,
  tipfacmanual smallint default 0 not null,
  tipfacempleado varchar (200),
  tipfacorigen varchar (200),
  tipfacdb varchar (200),
  tipfacprog varchar(60),
  constraint pk_anrtipofact primary key (tipfacnro)
);
CREATE UNIQUE INDEX tipofact_1 ON anrtipofact (tipfacnro);
CREATE INDEX tipfacdesabr_1 ON anrtipofact (tipfacdesabr);

CREATE TABLE anrfactor  
( 
  facnro integer IDENTITY (1, 1) not null,
  facdesabr varchar(30),
  facdesext varchar(200),
  facestado smallint default -1 not null,
  facpropor smallint default 0 not null,
  tipfacnro integer not null,
  constraint pk_anrfactor primary key (facnro) 
);
CREATE UNIQUE INDEX anrfactor_1 ON anrfactor (facnro);
CREATE INDEX facdesabr_1 ON anrfactor (facdesabr);
CREATE INDEX tipfacnro_2 ON anrfactor (tipfacnro);

CREATE TABLE anrfact_ori  
( 
  facnro integer not null,
  faccodorig integer not null,
  tipfacnro integer not null
);
CREATE UNIQUE INDEX anrfactori_1 ON anrfact_ori (facnro,faccodorig);
CREATE INDEX faccodorig_1 ON anrfact_ori (faccodorig);
CREATE INDEX tfacnro_1 ON anrfact_ori (tipfacnro);

CREATE TABLE anrcab
(
 anrcabnro integer IDENTITY (1, 1) not null,
 anrcabdesabr varchar (30),
 anrcabdesext varchar (200),
 anrcabfecdesde datetime ,
 anrcabfechasta datetime ,
 anrcabreal smallint default -1 not null,
 constraint pk_anrcab primary key (anrcabnro) 
);
CREATE UNIQUE INDEX anrcab_1 ON anrcab (anrcabnro);
CREATE INDEX anrfechas ON anrcab (anrcabfecdesde,anrcabfechasta);
CREATE INDEX anrcabdesabr_1 ON anrcab (anrcabdesabr);

CREATE TABLE anrcab_fact 
( 
 facnro integer not null,
 anrcabnro integer not null,
 nrocolum integer not null
);
CREATE UNIQUE  INDEX  fact_anrem ON anrcab_fact  (facnro, anrcabnro);
CREATE UNIQUE  INDEX  anrem_fact ON anrcab_fact  (anrcabnro,facnro);

CREATE TABLE anrrangofec
( 
 anrrangnro integer IDENTITY (1, 1) not null,
 anrcabnro integer not null,
 anrrangorden integer,
 anrrangdesabr varchar(30),
 anrrangfecdesde datetime,
 anrrangfechasta datetime,
 constraint pk_anrrangofec primary key (anrrangnro) 
);
CREATE UNIQUE INDEX anrrangofec  ON anrrangofec (anrrangnro);
CREATE INDEX anrrangdesabr ON anrrangofec (anrrangdesabr);
CREATE INDEX anrrfec ON anrrangofec (anrrangfecdesde,anrrangfechasta);
CREATE INDEX anrcab_3 ON anrrangofec (anrcabnro);
CREATE INDEX anrrangorden ON anrrangofec (anrcabnro,anrrangorden);

CREATE TABLE anrcubo 
(
 anrcabnro integer not null,
 facnro integer not null,
 anrcubvalor1 decimal(9,2),
 anrcubvalor2 decimal(9,2),
 anrcubvalor3 decimal(9,2),
 anrcubvalor4 decimal(9,2),
 anrcubvalor5 decimal(9,2),
 estrnro integer,
 tenro integer,
 ternro integer,
 tipnro integer,
 anrcubmanual smallint default 0 not null, 
 anrrangnro integer not null 
);
CREATE INDEX  cab_fact_1 ON anrcubo (anrcabnro ,facnro);
CREATE INDEX  anrcabcab ON anrcubo (anrcabnro);
CREATE INDEX  anrestr ON anrcubo (anrcabnro,facnro,tenro,estrnro);
CREATE INDEX  anrempleado ON anrcubo (anrcabnro,facnro,ternro);
CREATE INDEX  anrrango ON anrcubo (anrcabnro,facnro,anrrangnro);
CREATE UNIQUE INDEX anrcubo ON anrcubo (anrcabnro,facnro,tenro,
estrnro,ternro,anrrangnro);
CREATE INDEX  anrcubo_2 ON anrcubo (ternro,anrrangnro,facnro);

CREATE TABLE anrcab_agrup 
(
 anrcabnro integer not null,
 estrnro integer,
 tenro integer,
 anrcagorden integer not null,
 anrcagtotaliz smallint default -1 not null 
);
CREATE INDEX anrcab_agrup_1 ON anrcab_agrup (anrcabnro ,tenro,estrnro);

CREATE TABLE anrcab_filtro 
(
 anrcabnro integer not null,
 estrnro integer,
 tenro integer
);
CREATE INDEX anrcab_filtro_1 ON anrcab_filtro (anrcabnro ,tenro,estrnro);

CREATE TABLE anrcab_factot
( 
 facnro integer not null,
 anrcabnro integer not null,
 facnrotot integer not null
);
CREATE UNIQUE INDEX  anrcab_factot ON anrcab_factot (anrcabnro,
facnro, facnrotot);
CREATE INDEX factot_1 ON anrcab_factot (facnro);
CREATE INDEX factot_2 ON anrcab_factot (facnrotot);


alter table anrfactor add constraint fk_factipofact foreign 
    key (tipfacnro) references anrtipofact;

alter table anrcab_fact add constraint fk_cabfaccab foreign 
    key (anrcabnro)references anrcab;

alter table anrcab_fact add constraint fk_cabfacfac foreign 
    key (facnro) references anrfactor;

alter table anrrangofec add constraint fk_ranfeccab foreign 
    key (anrcabnro) references anrcab ;

alter table anrcubo add constraint fk_cubocab foreign 
    key (anrcabnro) references anrcab;

alter table anrcubo add constraint fk_cubofac foreign 
    key (facnro) references anrfactor;
  
alter table anrcubo add constraint fk_cuboestr foreign 
   key (estrnro) references estructura;

alter table anrcubo add constraint fk_cubote foreign 
    key (tenro) references tipoestructura;

alter table anrcubo add constraint fk_cuboter foreign 
    key (ternro) references tercero 
    ;

alter table anrcab_agrup add constraint fk_agrucab foreign 
    key (anrcabnro) references anrcab ;

alter table anrcab_agrup add constraint fk_agruestr foreign 
    key (estrnro) references estructura     ;

alter table anrcab_agrup add constraint fk_agrute foreign 
    key (tenro) references tipoestructura     ;

alter table anrcab_filtro add constraint fk_filcab foreign 
    key (anrcabnro) references anrcab      ;

alter table anrcab_filtro add constraint fk_filestr foreign 
    key (estrnro) references estructura  ;

alter table anrcab_filtro add constraint fk_filte foreign 
    key (tenro) references tipoestructura    ;

alter table anrcab_factot add constraint fk_totcab foreign 
    key (anrcabnro) references anrcab  ;

alter table anrcab_factot add constraint fk_totfac foreign 
    key (facnro) references anrfactor    ;

 