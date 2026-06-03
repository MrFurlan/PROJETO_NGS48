CREATE TABLE [dbo].[TelaSelecao](
	[Processo_Id] [nvarchar](50) NOT NULL,
	[TextoConsulta] [varchar](8000) NOT NULL,
	[TipoSelecao_Id] [varchar](3) NOT NULL,
 CONSTRAINT [PK_TelaSelecao] PRIMARY KEY CLUSTERED 
(
	[Processo_Id] ASC,
	[TipoSelecao_Id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

CREATE TABLE [dbo].[TelaSelecaoParametros](
	[Processo_Id] [nvarchar](50) NOT NULL,
	[Parametro_Id] [varchar](50) NOT NULL,
	[TipoSelecao_Id] [varchar](3) NOT NULL,
	[TipoValor] [varchar](50) NOT NULL,
	[Tamanho] [int] NOT NULL,
	[Decimal] [int] NOT NULL,
 CONSTRAINT [PK_TelaSelecaoParametros] PRIMARY KEY CLUSTERED 
(
	[Processo_Id] ASC,
	[Parametro_Id] ASC,
	[TipoSelecao_Id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[TelaSelecao]  WITH CHECK ADD  CONSTRAINT [FK_TelaSelecao_Telas] FOREIGN KEY([Processo_Id])
REFERENCES [dbo].[Telas] ([Processo_Id])
GO

ALTER TABLE [dbo].[TelaSelecao] CHECK CONSTRAINT [FK_TelaSelecao_Telas]
GO

ALTER TABLE [dbo].[TelaSelecao]  WITH CHECK ADD  CONSTRAINT [FK_TelaSelecao_TiposSelecao] FOREIGN KEY([TipoSelecao_Id])
REFERENCES [dbo].[TiposSelecao] ([TipoSelecao_Id])
GO

ALTER TABLE [dbo].[TelaSelecao] CHECK CONSTRAINT [FK_TelaSelecao_TiposSelecao]
GO

ALTER TABLE [dbo].[TelaSelecaoParametros]  WITH CHECK ADD  CONSTRAINT [FK_TelaSelecaoParametros_TelaSelecao1] FOREIGN KEY([Processo_Id], [TipoSelecao_Id])
REFERENCES [dbo].[TelaSelecao] ([Processo_Id], [TipoSelecao_Id])
GO

ALTER TABLE [dbo].[TelaSelecaoParametros] CHECK CONSTRAINT [FK_TelaSelecaoParametros_TelaSelecao1]
GO

IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_ListaSelecao_Telas]') AND parent_object_id = OBJECT_ID(N'[dbo].[ListaSelecao]'))
ALTER TABLE [dbo].[ListaSelecao] DROP CONSTRAINT [FK_ListaSelecao_Telas]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaSelecao]') AND type in (N'U'))
DROP TABLE [dbo].[ListaSelecao]
GO

ALTER TABLE [dbo].[TelaSelecao]  WITH CHECK ADD  CONSTRAINT [FK_TelaSelecao_Telas] FOREIGN KEY([Processo_Id])
REFERENCES [dbo].[Telas] ([Processo_Id])
GO

ALTER TABLE [dbo].[TelaSelecao] CHECK CONSTRAINT [FK_TelaSelecao_Telas]
GO

ALTER TABLE [dbo].[TelaSelecao]  WITH CHECK ADD  CONSTRAINT [FK_TelaSelecao_TiposSelecao] FOREIGN KEY([TipoSelecao_Id])
REFERENCES [dbo].[TiposSelecao] ([TipoSelecao_Id])
GO

ALTER TABLE [dbo].[TelaSelecao] CHECK CONSTRAINT [FK_TelaSelecao_TiposSelecao]
GO

INSERT INTO TelaSelecao VALUES ('Estados', 'SELECT Estados.Estado_Id, Estados.Descricao FROM Estados', 'SEL')
INSERT INTO TelaSelecao VALUES ('Estados', 'SELECT Estados.Estado_Id, Estados.Descricao FROM Estados WHERE Estados.Estado_Id = @Estado_Id', 'SFR')