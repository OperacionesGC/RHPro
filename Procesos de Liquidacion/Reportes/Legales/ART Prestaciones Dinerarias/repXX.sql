CREATE TABLE rep06 (
	repnro int IDENTITY (1, 1) NOT NULL ,
	pliqnro int NOT NULL ,
	pronro int NOT NULL ,
	empresa int NOT NULL ,
	proaprob int NOT NULL ,
	bpronro int NOT NULL ,
	Fecha datetime NOT NULL ,
	Hora varchar (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	iduser varchar (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
)

