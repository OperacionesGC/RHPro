CREATE TABLE  sotipoaccidente ( 
 tipoaccnro  INTEGER IDENTITY (1, 1) not null, 
tipoaccdabr varchar (30) not null, tipoaccdext varchar (255),
 constraint pk_tipoaccidente primary key (tipoaccnro) 
 );CREATE UNIQUE INDEX  tipoacc_1 ON sotipoaccidente (tipoaccnro);CREATE UNIQUE INDEX  tipoacc_2 ON sotipoaccidente (tipoaccdabr);


CREATE TABLE  socausaacc 
( 
 cauaccnro INTEGER IDENTITY (1, 1) not null, cauaccdesabr varchar (30)  not null, cauaccdesext varchar (255),
 constraint pk_causaacc primary key (cauaccnro)
);CREATE UNIQUE  INDEX  causaacc_1 ON socausaacc (cauaccnro);CREATE UNIQUE  INDEX  causaacc_2 ON socausaacc (cauaccdesabr);

CREATE TABLE  soaccid_visita 
( 
 visitamed integer not null,
 accnro integer not null
);CREATE UNIQUE INDEX  accvisita_1 ON soaccid_visita (visitamed,accnro);
CREATE INDEX  accvisita_2 ON soaccid_visita (visitamed);
CREATE INDEX  accvisita_3 ON soaccid_visita (accnro);




CREATE TABLE  soaccidente 
( 
 accnro INTEGER IDENTITY (1, 1) not null, accdescext varchar (255), accfecha datetime not null, acchora varchar , accpoliza varchar (10), accincapacidad varchar (255), accsecuelas varchar (255), actinsegnro integer not null, empleado integer not null, cauaccnro integer not null, tipoaccnro integer not null, condinnro integer not null, factaccnro integer not null, formaaccnro integer not null, zonacunro integer not null, natulenro integer not null, agenmatnro integer not null, acclugar varchar (30), accdomicilio varchar (30), acctestigos varchar (30), acctarearealiz varchar (30), accocupregular varchar (30), accnivelconoc varchar (30), accdenunciante varchar (30), acctareahab smallint, accentrenamiento smallint , acccapaci smallint,
 constraint ck_acctareahab check ((acctareahab = 0 ) OR (acctareahab = -1 ) ),
 constraint ck_accentren check ((accentrenamiento = 0 ) OR (accentrenamiento = -1 ) ) ,
 constraint ck_acccapaci check ((acccapaci = 0 ) OR (acccapaci = -1 ) ) ,
 constraint pk_accidente primary key (accnro)
);CREATE UNIQUE INDEX accidente_1 ON soaccidente (accnro);
CREATE INDEX  accidente_2 ON soaccidente (empleado);
CREATE INDEX  accidente_3 ON soaccidente (accfecha,acchora);CREATE INDEX  accidente_4 ON soaccidente (cauaccnro);CREATE INDEX  accidente_5 ON soaccidente (tipoaccnro);
CREATE INDEX  accidente_6 ON soaccidente (actinsegnro);
CREATE INDEX  accidente_7 ON soaccidente (condinnro);CREATE INDEX  accidente_8 ON soaccidente (factaccnro);CREATE INDEX  accidente_9 ON soaccidente (formaaccnro);CREATE INDEX  accidente_10 ON soaccidente (zonacunro);CREATE INDEX  accidente_11 ON soaccidente (natulenro);CREATE INDEX  accidente_12 ON soaccidente (agenmatnro);
CREATE TABLE  soactosinseg 
( actinsegnro INTEGER IDENTITY (1, 1) not null,  actinsegdesabr varchar (30) not null,  actinsegdesext varchar (255),
  constraint pk_actosinseg primary key (actinsegnro)
);CREATE UNIQUE INDEX  actinsegnro ON soactosinseg (actinsegnro);
CREATE UNIQUE INDEX  actinsegdesabr ON soactosinseg (actinsegdesabr);

CREATE TABLE  soagenmat 
( agenmatnro integer identity (1,1) not null ,  agenmatdesabr varchar (30) not null,  agenmatdesext varchar (200),
  constraint pk_agenmat primary key (agenmatnro)
);CREATE UNIQUE  INDEX  agenmatnro ON soagenmat (agenmatnro  );CREATE UNIQUE  INDEX  agenmatdesabr ON soagenmat (agenmatdesabr);


CREATE TABLE  socondinsegura 
( 
 condinnro INTEGER IDENTITY (1,1) not null, condindesabr varchar (30) not null, condindesext varchar (200),
 constraint pk_condinseg primary key (condinnro)
);CREATE UNIQUE  INDEX  condinnro ON socondinsegura (condinnro);CREATE UNIQUE  INDEX  conindesabr ON socondinsegura (condindesabr);


CREATE TABLE  sofactaccidente 
( 
 factaccnro INTEGER IDENTITY (1,1) not null, factaccdesabr varchar (30) not null, factaccdesext varchar (200),
 constraint pk_factaccidente primary key (factaccnro)
);CREATE UNIQUE  INDEX  Factaccnro ON sofactaccidente (factaccnro);CREATE UNIQUE  INDEX  factaccdesabr ON sofactaccidente (factaccdesabr);


CREATE TABLE  somedico 
(
 ternro integer not null, horatendesde varchar , horatenhasta varchar , medubicacion smallint, medcosto decimal(9,2), medconc varchar (200), medespec varchar (200),
 constraint ck_medubicacion check ((medubicacion = 0 ) OR (medubicacion = -1 ) ));CREATE UNIQUE  INDEX  medico_1 ON somedico (ternro);


CREATE TABLE sopatologias 
( 
 patologianro INTEGER IDENTITY (1,1) not null, patologiadesabr varchar (30) not null, patologiadesext varchar (255), patologiacodext varchar (20),
 constraint pk_patologia primary key (patologianro)
);CREATE UNIQUE  INDEX  patologianro ON sopatologias (patologianro);CREATE UNIQUE  INDEX  patologiadesabr ON sopatologias (patologiadesabr);
CREATE INDEX  patologiacodext ON sopatologias (patologiacodext);
CREATE TABLE  sopatol_visitas 
( 
 visitamed integer not null, patologianro integer not null
);CREATE UNIQUE INDEX  patolvisita_1 ON sopatol_visitas (visitamed  ,patologianro );
CREATE INDEX  patolvisita_2 ON sopatol_visitas (visitamed);
CREATE INDEX  patolvisita_3 ON sopatol_visitas (patologianro );

CREATE TABLE  soservmedico 
( 
 ternro integer not null, servmedcosto decimal (9,2), servmeddesc varchar (255), servmedresp varchar (30), servmedobs varchar (255));CREATE UNIQUE  INDEX  servmedico_1 ON soservmedico (ternro);

CREATE TABLE  sotipvisitas ( 
 tipvinro INTEGER IDENTITY (1,1) not null , 
 tipvidesabr character (30), 
 tipvidesext character (255), 
 tipvisis smallint,
 constraint ck_tipvisis check ((tipvisis = 0 ) OR (tipvisis= -1 ) ),
 constraint pk_tipovis primary key (tipvinro) 

);
CREATE UNIQUE  INDEX  tiponro       ON sotipvisitas (tipvinro);CREATE UNIQUE  INDEX  tipodesabr ON tipvisitas (tipvidesabr);

CREATE TABLE sovisitamedica 
( 
 vismednro INTEGER IDENTITY (1,1) not null , vismeddesc varchar (255), vismedubicacion smallint , vismedfecha datetime , vismedhora varchar , vismeddiag varchar (255), medico integer , ternro integer not null, servmedico integer , vismedcosto decimal (9,2), tipvinro integer not null, vismedempdir varchar (255),
 constraint ck_vismedubi check ((vismedubicacion = 0 ) OR (vismedubicacion = -1 ) ),
 constraint pk_visitamedica primary key (vismednro)
);CREATE UNIQUE  INDEX  vismednro_1 ON sovisitamedica (vismednro);
CREATE INDEX  vismednro_2 ON sovisitamedica (ternro);
CREATE INDEX  vismednro_3 ON sovisitamedica (tipvinro);
CREATE INDEX  vismednro_4 ON sovisitamedica (servmedico);
CREATE INDEX  vismednro_5 ON sovisitamedica (medico);
CREATE TABLE  somedi_visita 
( 
 medinro integer not null, vismednro integer not null, mvrece smallint not null, mvcant integer,
 constraint ck_mvrece check ((mvrece = 0 ) OR (mvrece = -1 ) )
);CREATE UNIQUE  INDEX  medi_visita_1 ON somedi_visita (medinro,vismednro);
CREATE UNIQUE  INDEX  medi_visita_2 ON somedi_visita (vismednro);
CREATE UNIQUE  INDEX  medi_visita_3 ON somedi_visita (medinro);

CREATE TABLE  somedicamentos 
( 
 medinro INTEGER IDENTITY (1,1) not null , medicodext varchar (10), medidesabr varchar (30), medidesext varchar (255),
 constraint pk_medicamentos primary key (medinro)
);CREATE UNIQUE  INDEX  medicamento ON somedicamentos (medinro);
CREATE INDEX  medicamento_1 ON somedicamentos (medicodext);
CREATE INDEX  medicamento_2 ON somedicamentos (medidesabr);

CREATE TABLE  sofichamedica 
( 
 fmednro INTEGER IDENTITY (1, 1) not null, fmedavisa varchar (50), fmedtel varchar (15), fmednotas varchar (255), fmedsafety varchar (50), ternro integer , fmedfecha datetime , fmeddesabr varchar (30), fmednotadisc varchar (255), fmedgrupo varchar (8), fmedfactor varchar (8), fmednotapat varchar (255), fmednotamed varchar (255), fmeddiscap smallint , fmedporcen decimal(5,2),
 constraint ck_fmeddiscap check ((fmeddiscap = 0 ) OR (fmeddiscap = -1 ) ),
 constraint pk_fichamedica primary key (fmednro)
);CREATE UNIQUE  INDEX  ficham_1 ON sofichamedica (fmednro);CREATE INDEX  ficham_2 ON sofichamedica (fmedfecha,ternro);CREATE INDEX  ficham_3 ON sofichamedica (fmedfecha);CREATE INDEX  ficham_4 ON sofichamedica (ternro,fmedfecha);CREATE INDEX  ficham_5 ON sofichamedica (ternro);

CREATE TABLE  sofmed_med 
( 
  fmednro integer not null,  medinro integer not null
);CREATE INDEX  fmedm_1 ON sofmed_med (fmednro  );CREATE INDEX  fmedm_2 ON sofmed_med (fmednro,medinro);CREATE INDEX  fmedm_3 ON sofmed_med (medinro);CREATE INDEX  fmedm_4 ON sofmed_med (medinro,fmednro);

CREATE TABLE  sofmed_pat 
( 
 fmednro integer not null, patologianro integer not null
);CREATE INDEX  fmedp_1 ON sofmed_pat (fmednro);CREATE INDEX  fmedp_2 ON sofmed_pat (fmednro,patologianro);CREATE INDEX  fmedp_3 ON sofmed_pat (patologianro,fmednro);CREATE INDEX  fmedp_4 ON sofmed_pat (patologianro);

CREATE TABLE  soformaaccidente 
( 
 formaaccnro INTEGER IDENTITY (1, 1) not null, formaaccdesabr varchar (30) not null, formaaccdesext varchar (200),
 constraint pk_formaaccidente primary key (formaaccnro)
);CREATE UNIQUE  INDEX  formaacc_1 ON soformaaccidente (formaaccnro);
CREATE UNIQUE  INDEX  formaacc_2 ON soformaaccidente (formaaccdesabr);

CREATE TABLE  sonatulesion 
( 
 natulenro INTEGER IDENTITY (1, 1) not null, natuledesabr varchar (30) not null, natuledesext varchar (200),
 constraint pk_natulesion primary key (natulenro)
);CREATE UNIQUE  INDEX  natule_1 ON sonatulesion (natulenro);CREATE UNIQUE  INDEX  natule_2 ON sonatulesion (natuledesabr);

CREATE TABLE  sozonacuerpo 
( 
 zonacunro INTEGER IDENTITY (1, 1) not null, zonacudesabr varchar (30) not null, zonacudesext varchar (200)
);CREATE UNIQUE  INDEX  zonacu_1 ON sozonacuerpo (zonacunro);CREATE UNIQUE  INDEX  zonacu_2 ON sozonacuerpo (zonacudesabr);
