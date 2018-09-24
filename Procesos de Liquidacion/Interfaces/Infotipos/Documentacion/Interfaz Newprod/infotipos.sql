CREATE TABLE infotipos (
	inftipnro int IDENTITY (1, 1) NOT NULL ,
	inftipcod varchar (10)  NULL ,
	inftipdesabr varchar (60)  NULL ,
	inftipdesext varchar (100)  NULL ,
	inftiporden smallint NULL 
)

CREATE TABLE infotipos_mapeo (
	tablaref varchar (10)  NULL ,
	codinterno varchar (10)  NULL ,
	codexterno varchar (10)  NULL 
)

