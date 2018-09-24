if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rep_asp_01]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[rep_asp_01]
GO

CREATE TABLE [dbo].[rep_asp_01] (
	[bprcnro] [bigint] NOT NULL ,
	[repnro] [bigint] NOT NULL ,
	[ternro] [bigint] NOT NULL ,
	[fecha] [datetime] NOT NULL ,
	[causa] [varchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[descripcion] [varchar] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[horas] [decimal](18, 2) NULL ,
	[replinnro] [bigint] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

