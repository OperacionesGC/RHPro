if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[inter_err]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[inter_err]
GO

CREATE TABLE [dbo].[inter_err] (
	[crpnnro] [int] NOT NULL ,
	[inerrnro] [int] NOT NULL ,
	[nrolinea] [int] NOT NULL ,
	[campnro] [int] NOT NULL 
) ON [PRIMARY]
GO

