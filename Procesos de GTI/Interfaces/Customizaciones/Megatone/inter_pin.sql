if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[inter_pin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[inter_pin]
GO

CREATE TABLE [dbo].[inter_pin] (
	[crpnnro] [int] IDENTITY (1, 1) NOT NULL ,
	[modnro] [int] NOT NULL ,
	[crpnarchivo] [varchar] (60) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[crpnregleidos] [int] NOT NULL ,
	[crpnregerr] [int] NOT NULL ,
	[crpnfecha] [datetime] NOT NULL ,
	[crpndesc] [varchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[crpnestado] [varchar] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[bpronro] [int] NOT NULL 
) ON [PRIMARY]
GO

