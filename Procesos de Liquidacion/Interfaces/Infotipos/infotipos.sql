CREATE TABLE infotipos(
   inftipnro    int identity(1,1) not null,
   inftipcod    varchar(10),
   inftipdesabr varchar(60),
   inftipdesext varchar(100),
   inftiporden  smallint
)

CREATE TABLE modelo_infotipos(
   modnro int,
   inftipnro int
)

CREATE TABLE infotipos_mapeo(
   tablaref   varchar(10),
   codinterno varchar(10),
   codexterno varchar(10)
   
)