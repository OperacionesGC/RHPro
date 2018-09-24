CREATE TABLE rep08 (
	repnro int IDENTITY (1, 1) NOT NULL ,
	pliqnro int NOT NULL ,
	pronro int NOT NULL ,
	empresa int NOT NULL ,
	proaprob int NOT NULL ,
	bpronro int NOT NULL ,
	Fecha datetime NOT NULL ,
	Hora varchar (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	iduser varchar (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	concnro int NOT NULL ,
	agrupacion int NULL ,
	agrupacion2 int NULL ,
	agrupacion3 int NULL ,
	total_liquidado decimal(19, 4) NOT NULL ,
	cant_liquidado decimal(19, 4) NOT NULL ,
	emp_liquidado int NOT NULL ,
	asigfam decimal(19, 4) NOT NULL 
)

