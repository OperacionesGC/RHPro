CREATE TABLE rep41 (
	repnro int IDENTITY (1, 1) NOT NULL ,
	pliqnro int NOT NULL ,
	pronro int NOT NULL ,
	empresa int NOT NULL ,
	proaprob int NOT NULL ,
	bpronro int NOT NULL ,
	Fecha datetime NOT NULL ,
	Hora varchar (10)  NOT NULL ,
	iduser varchar (20)  NOT NULL ,
	billdes varchar(20) NOT null,
	billcan decimal(19,4) NOT null,
	moncod int not null,
	mondes varchar(20) not null,
	moncan int not null,
	monnro int not null,
	billcod decimal(19,4) not null
)

