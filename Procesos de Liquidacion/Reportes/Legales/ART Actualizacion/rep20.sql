if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rep20]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[rep20]
GO

CREATE TABLE [dbo].[rep20] (
	[repnro] [int] IDENTITY (1, 1) NOT NULL ,
	[bpronro] [int] NOT NULL ,
	[empresa] [int] NULL ,
	[fecha] [datetime] NOT NULL ,
	[hora] [varchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[iduser] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[ternro] [int] NOT NULL ,
	[ternom] [varchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[ternom2] [varchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[terape] [varchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[terape2] [varchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[terfecnac] [datetime] NULL ,
	[tersex] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[terestciv] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[calle] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[nro] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[torre] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[piso] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[oficdepto] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[codigopostal] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[telnro] [varchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,
	[tidsigla] [varchar] (5) COLLATE Modern_Spanish_CI_AS NULL ,
	[nrodoc] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[locdesc] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[provdesc] [varchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[nacionalidad] [int] NULL ,
	[sucursal] [int] NULL ,
	[cuil] [varchar] (13) COLLATE Modern_Spanish_CI_AS NULL ,
	[contratacion] [smallint] NULL ,
	[ingreso] [datetime] NULL ,
	[legajo] [int] NULL ,
	[remuneracion] [decimal](19, 4) NULL ,
	[columna1] [decimal](19, 4) NULL ,
	[columna2] [decimal](19, 4) NULL ,
	[columna3] [decimal](19, 4) NULL ,
	[columna4] [decimal](19, 4) NULL ,
	[columna5] [decimal](19, 4) NULL 
) ON [PRIMARY]
GO

