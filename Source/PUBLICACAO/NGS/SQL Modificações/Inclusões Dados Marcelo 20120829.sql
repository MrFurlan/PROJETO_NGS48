INSERT INTO Telas VALUES ('Estados',	'Estados',	0,	1,	1)

INSERT INTO TiposCampo VALUES ('CalendarBox', 		0,	0,	0)
INSERT INTO TiposCampo VALUES ('CheckBox',			0,	0,	0)
INSERT INTO TiposCampo VALUES ('DescriptionSearch',	0,	1,	1)
INSERT INTO TiposCampo VALUES ('DropDownList',		1,	0,	0)
INSERT INTO TiposCampo VALUES ('NumberBox',			0,	0,	0)
INSERT INTO TiposCampo VALUES ('OptionList',		1,	0,	0)
INSERT INTO TiposCampo VALUES ('TextBox',			0,	0,	0)

INSERT INTO TelaCampos VALUES ('Estados_txtEstadoId', 'Estados', 'Sigla', 'Estado_Id', 2, 1, 'TextBox', 1, 1, 'String')
INSERT INTO TelaCampos VALUES ('Estados_txtEstadoNome', 'Estados', 'Nome', 'Descricao', 50, 1, 'TextBox', 2, 0, 'String')

INSERT INTO ListaColunas VALUES ('Estados', 'Estados_clhEstadoId', 'Sigla', 'Estados', 'Estado_Id', 100, 1, 1, 1, 1, 'String', 'Estados_txtEstadoId')
INSERT INTO ListaColunas VALUES ('Estados', 'Estados_clhEstadoNome', 'Nome', 'Estados', 'Descricao', 400, 1, 1, 2, 0, 'String', 'Estados_txtEstadoNome')

--Executar essas linhas antes de adicionar valores no ListaColunas
ALTER TABLE ListaColunas ADD [Campo_Id] [varchar](200) NULL

ALTER TABLE [dbo].[ListaColunas]  WITH CHECK ADD  CONSTRAINT [FK_ListaColunas_TelaCampos] FOREIGN KEY([Campo_Id])
REFERENCES [dbo].[TelaCampos] ([Campo_Id])
GO

ALTER TABLE [dbo].[ListaColunas] CHECK CONSTRAINT [FK_ListaColunas_TelaCampos]
GO