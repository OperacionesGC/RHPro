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
 constraint pk_anrrangofec primary key (anrrangonro) 
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


alter table anrfactor add constraint (foreign 
    key (tipfacnro) references anrtipofact constraint 
    fk_factipofact);

alter table anrcab_fact add constraint (foreign 
    key (anrcabnro) references anrcab constraint 
    fk_cabfaccab);

alter table anrcab_fact add constraint (foreign 
    key (facnro) references anrfactor constraint 
    fk_cabfacfac);

alter table anrrangofec add constraint (foreign 
    key (anrcabnro) references anrcab constraint 
    fk_ranfeccab);

alter table anrcubo add constraint (foreign 
    key (anrcabnro) references anrcab constraint 
    fk_cubocab);

alter table anrcubo add constraint (foreign 
    key (facnro) references anrfactor constraint 
    fk_cubofac);
  
alter table anrcubo add constraint (foreign 
    key (estrnro) references estructura constraint 
    fk_cuboestr);

alter table anrcubo add constraint (foreign 
    key (tenro) references tipoestructura constraint 
    fk_cubote);

alter table anrcubo add constraint (foreign 
    key (ternro) references tercero constraint 
    fk_cuboter);

alter table anrcab_agrup add constraint (foreign 
    key (anrcabnro) references anrcab  constraint 
    fk_agrucab);

alter table anrcab_agrup add constraint (foreign 
    key (estrnro) references estructura constraint 
    fk_agruestr);

alter table anrcab_agrup add constraint (foreign 
    key (tenro) references tipoestructura  constraint 
    fk_agrute);

alter table anrcab_filtro add constraint (foreign 
    key (anrcabnro) references anrcab  constraint 
    fk_filcab);

alter table anrcab_filtro add constraint (foreign 
    key (estrnro) references estructura constraint 
    fk_filestr);

alter table anrcab_filtro add constraint (foreign 
    key (tenro) references tipoestructura  constraint 
    fk_filte);

alter table anrcab_factot add constraint (foreign 
    key (anrcabnro) references anrcab  constraint 
    fk_totcab);

alter table anrcab_factot add constraint (foreign 
    key (facnro) references anrfactor constraint 
    fk_totfac);

 