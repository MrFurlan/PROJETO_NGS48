Imports System.IO
Imports System.Data
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeComissoesPedido
    Inherits BasePage

#Region "Variáveis"
    Dim Sql As String = ""
    Dim SqlArray As New ArrayList
#End Region

#Region "Métodos"

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe as datas.")
            Return False
        End If

        Return True
    End Function

    Private Sub CargaUnidadeDeNegocioEmpresa()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Function GetSql() As String
        If FinanceiroNovo Then

            Sql &= "   SELECT " & vbCrLf &
                   "           	Empresa.Reduzido AS RedEmpresa,                                                                        " & vbCrLf &
                   "   			Comissoes.Empresa_Id AS CNPJEmpresa,                                                                   " & vbCrLf &
                   "   			Comissoes.EndEmpresa_Id AS EndEmpresa,                                                                 " & vbCrLf &
                   "   			Empresa.Nome AS NomeEmpresa,                                                                           " & vbCrLf &
                   "   			Empresa.Cidade AS CidadeEmpresa,                                                                       " & vbCrLf &
                   "   			Empresa.Estado AS UFEmpresa,                                                                           " & vbCrLf &
                   "   			Comissoes.Representante_Id AS CNPJRepresentante,                                                       " & vbCrLf &
                   "   			Comissoes.EndRepresentante_Id AS EndRepresentante,                                                     " & vbCrLf &
                   "   			Representante.Nome AS NomeRepresentante,                                                               " & vbCrLf &
                   "   			Representante.Cidade AS CidadeRepresentante,                                                           " & vbCrLf &
                   "   			Representante.Estado AS UFRepresentante,                                                               " & vbCrLf &
                   "   			Clientes.Cliente_Id AS CnpjCliente,                                                                    " & vbCrLf &
                   "   			Clientes.Nome AS NomeDoCliente,                                                                        " & vbCrLf &
                   "   			Clientes.Cidade AS CidadeDoCliente,                                                                    " & vbCrLf &
                   "   			Comissoes.Pedido_Id AS Pedido,                                                                         " & vbCrLf &
                   "   			Pedidos.DataPedido,                                                                                    " & vbCrLf &
                   "   			SUM(NotasFiscaisXItens.QuantidadeFiscal) AS QuantidadeFaturada,                                        " & vbCrLf &
                   "   			SUM(NotasFiscaisXItens.Valor) AS ValorFaturado,                                                        " & vbCrLf &
                   "   			Comissoes.Percentual,                                                                                  " & vbCrLf &
                   "   			Case                                                                                                   " & vbCrLf &
                   "   			when Pedidos.Operacao <> 36                                                                            " & vbCrLf &
                   "   				then SUM(NotasFiscaisXItens.Valor) * Comissoes.Percentual / 100                                    " & vbCrLf &
                   "   				else ((SUM(NotasFiscaisXItens.QuantidadeFiscal / 1000) * Comissoes.Percentual) * Cotacoes.Indice)  " & vbCrLf &
                   "   						* (((case                                                                                  " & vbCrLf &
                   "   							 when M.Classificacao = 'O'                                                            " & vbCrLf &
                   "   									then VLDOCUMENTO.Valoroficial                                                  " & vbCrLf &
                   "   									else VLDOCUMENTO.ValorMoeda                                                    " & vbCrLf &
                   "   							 end --GET VALOR DOCUMENTO                                                             " & vbCrLf &
                   "   						/ SUM(NotasFiscaisXItens.Valor) * 100)) / 100)                                             " & vbCrLf &
                   "   			End  AS Comissao,                                                                                      " & vbCrLf &
                   "   			CR.Titulo_Id AS Titulo,                                                                                " & vbCrLf &
                   "   			CR.DataBaixa AS DataTitulo,                                                                            " & vbCrLf &
                   "   			CASE WHEN CR.Provisao = 1                                                                              " & vbCrLf &
                   "   				THEN case                                                                                          " & vbCrLf &
                   "   						when M.Classificacao = 'O'                                                                 " & vbCrLf &
                   "   						then VLDOCUMENTO.Valoroficial                                                              " & vbCrLf &
                   "   						else VLDOCUMENTO.ValorMoeda                                                                " & vbCrLf &
                   "   					 end --GET VALOR DOCUMENTO                                                                     " & vbCrLf &
                   "   				ELSE 0                                                                                             " & vbCrLf &
                   "   			END AS ValorPago,                                                                                      " & vbCrLf &
                   "   			CASE WHEN CR.Provisao = 1                                                                              " & vbCrLf &
                   "   				THEN case                                                                                          " & vbCrLf &
                   "   						when M.Classificacao = 'O'                                                                 " & vbCrLf &
                   "   						then VLDOCUMENTO.Valoroficial                                                              " & vbCrLf &
                   "   						else VLDOCUMENTO.ValorMoeda                                                                " & vbCrLf &
                   "   					 end --GET VALOR DOCUMENTO                                                                     " & vbCrLf &
                   "   				ELSE 0                                                                                             " & vbCrLf &
                   "   			END * Comissoes.Percentual / 100 AS ValorLiberado,                                                     " & vbCrLf &
                   "                       Pedidos.Operacao,                                                                              " & vbCrLf &
                   "                       Cotacoes.Indice,                                                                               " & vbCrLf &
                   "                       CR.RecPag as Conta                                                                             " & vbCrLf &
                   "                       FROM Comissoes                                                                                 " & vbCrLf &
                   "   	   LEFT JOIN Titulos CR                                                                                        " & vbCrLf &
                   "   	     ON Comissoes.Empresa_Id    = CR.Empresa                                                                   " & vbCrLf &
                   "           AND Comissoes.EndEmpresa_Id = CR.EndEmpresa                                                                " & vbCrLf &
                   "           AND Comissoes.Pedido_Id     = CR.Pedido                                                                    " & vbCrLf &
                   "   	  INNER JOIN TitulosxContaContabil VLDOCUMENTO                                                                 " & vbCrLf &
                   "   		 on CR.Titulo_Id           = VLDOCUMENTO.Titulo_Id                                                         " & vbCrLf &
                   "   		and CR.ContaContabilCliFor = VLDOCUMENTO.Conta_Id                                                          " & vbCrLf &
                   "   		and VLDOCUMENTO.DC_Id      = case                                                                          " & vbCrLf &
                   "   									   when CR.RecPag = 'R'                                                        " & vbCrLf &
                   "   									     then 'C'                                                                  " & vbCrLf &
                   "   										 else 'D'                                                                  " & vbCrLf &
                   "                       End                                                                                            " & vbCrLf &
                   "   	  INNER JOIN Cotacoes                                                                                          " & vbCrLf &
                   "   		 ON CR.DataBaixa = Cotacoes.Data_Id                                                                        " & vbCrLf &
                   "           AND CR.Indexador = Cotacoes.Indexador_Id                                                                   " & vbCrLf &
                   "         INNER JOIN Pedidos                                                                                           " & vbCrLf &
                   "            ON Comissoes.Pedido_Id     = Pedidos.Pedido_Id                                                            " & vbCrLf &
                   "           AND Comissoes.Empresa_Id    = Pedidos.Empresa_Id                                                           " & vbCrLf &
                   "           AND Comissoes.EndEmpresa_Id = Pedidos.EndEmpresa_Id                                                        " & vbCrLf &
                   "          LEFT JOIN Clientes                                                                                          " & vbCrLf &
                   "            ON Clientes.Cliente_Id  = Pedidos.Cliente                                                                 " & vbCrLf &
                   "           AND Clientes.Endereco_Id = Pedidos.EndCliente                                                              " & vbCrLf &
                   "         INNER JOIN NotasFiscaisXItens                                                                                " & vbCrLf &
                   "            ON NotasFiscaisXItens.Pedido = Pedidos.Pedido_Id                                                          " & vbCrLf &
                   "          LEFT JOIN Clientes AS Representante                                                                         " & vbCrLf &
                   "            ON Comissoes.Representante_Id    = Representante.Cliente_Id                                               " & vbCrLf &
                   "           AND Comissoes.EndRepresentante_Id = Representante.Endereco_Id                                              " & vbCrLf &
                   "          LEFT JOIN Clientes AS Empresa                                                                               " & vbCrLf &
                   "            ON Comissoes.Empresa_Id    = Empresa.Cliente_Id                                                           " & vbCrLf &
                   "           AND Comissoes.EndEmpresa_Id = Empresa.Endereco_Id                                                          " & vbCrLf &
                   "                                                                                                                      " & vbCrLf &
                   "           INNER JOIN Moedas M                                                                                        " & vbCrLf &
                   "   			on M.Moeda_id = CR.Moeda                                                                               " & vbCrLf &
                   "                                                                                                                      " & vbCrLf &
                   "                       WHERE Pedidos.Situacao = 1                                                                     " & vbCrLf &
                   "       and isnull(CR.RegistroMestre,0) <> CR.Titulo_Id                                                                " & vbCrLf &
                   "       And CR.Baixa  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf &
                   "       and CR.RecPag = 'R'                                                                                            " & vbCrLf &
                   "    and Comissoes.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf &
                   "    and Comissoes.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

            If ddlRepresentante.SelectedIndex > 0 Then
                Sql &= "        And Comissoes.Representante_Id = '" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            Sql &= "   GROUP BY Empresa.Reduzido,                                                                                            " & vbCrLf &
                   "   			Comissoes.Empresa_Id ,                                                                                       " & vbCrLf &
                   "   			Comissoes.EndEmpresa_Id ,                                                                                    " & vbCrLf &
                   "   			Empresa.Nome ,                                                                                               " & vbCrLf &
                   "   			Empresa.Cidade ,                                                                                             " & vbCrLf &
                   "   			Empresa.Estado ,                                                                                             " & vbCrLf &
                   "   			Comissoes.Representante_Id ,                                                                                 " & vbCrLf &
                   "   			Comissoes.EndRepresentante_Id ,                                                                              " & vbCrLf &
                   "   			Representante.Nome,                                                                                          " & vbCrLf &
                   "   			Representante.Cidade,                                                                                        " & vbCrLf &
                   "   			Representante.Estado,                                                                                        " & vbCrLf &
                   "   			Comissoes.Pedido_Id,                                                                                         " & vbCrLf &
                   "   			Comissoes.Percentual,                                                                                        " & vbCrLf &
                   "   			Pedidos.DataPedido,                                                                                          " & vbCrLf &
                   "   			CR.Titulo_Id,                                                                                                " & vbCrLf &
                   "               CR.RecPag,                                                                                                " & vbCrLf &
                   "   			CR.DataBaixa,			                                                                                     " & vbCrLf &
                   "   			Clientes.Cliente_Id,                                                                                         " & vbCrLf &
                   "   			Clientes.Nome,                                                                                               " & vbCrLf &
                   "   			Clientes.Cidade,                                                                                             " & vbCrLf &
                   "   			CR.Provisao,			                                                                                     " & vbCrLf &
                   "   			Pedidos.Operacao,                                                                                            " & vbCrLf &
                   "   			Cotacoes.Indice,                                                                                             " & vbCrLf &
                   "   			M.Classificacao,                                                                                             " & vbCrLf &
                   "   			VLDOCUMENTO.Valoroficial,                                                                                    " & vbCrLf &
                   "   			VLDOCUMENTO.ValorMoeda                                                                                       " & vbCrLf &
                   "                                                                                                                         " & vbCrLf &
                   "     Having   ISNULL(SUM(CASE                                                                                            " & vbCrLf &
                   "   						WHEN CR.Provisao in(1,2)                                                                         " & vbCrLf &
                   "   							THEN case                                                                                    " & vbCrLf &
                   "   									when M.Classificacao = 'O'                                                           " & vbCrLf &
                   "   										then VLDOCUMENTO.Valoroficial                                                    " & vbCrLf &
                   "           Else  VLDOCUMENTO.ValorMoeda                                                                                  " & vbCrLf &
                   "               End                                                                                                       " & vbCrLf &
                   "   							ELSE 0                                                                                       " & vbCrLf &
                   "   						END), 0) <> 0                                                                                    " & vbCrLf &
                   "                                                                                                                         " & vbCrLf &
                   "               Union all                                                                                                 " & vbCrLf &
                   "                                                                                                                         " & vbCrLf &
                   "   	 SELECT                                                                                                              " & vbCrLf &
                   "           	Empresa.Reduzido AS RedEmpresa,                                                                              " & vbCrLf &
                   "   			Comissoes.Empresa_Id AS CNPJEmpresa,                                                                         " & vbCrLf &
                   "   			Comissoes.EndEmpresa_Id AS EndEmpresa,                                                                       " & vbCrLf &
                   "   			Empresa.Nome AS NomeEmpresa,                                                                                 " & vbCrLf &
                   "   			Empresa.Cidade AS CidadeEmpresa,                                                                             " & vbCrLf &
                   "   			Empresa.Estado AS UFEmpresa,                                                                                 " & vbCrLf &
                   "   			Comissoes.Representante_Id AS CNPJRepresentante,                                                             " & vbCrLf &
                   "   			Comissoes.EndRepresentante_Id AS EndRepresentante,                                                           " & vbCrLf &
                   "   			Representante.Nome AS NomeRepresentante,                                                                     " & vbCrLf &
                   "   			Representante.Cidade AS CidadeRepresentante,                                                                 " & vbCrLf &
                   "   			Representante.Estado AS UFRepresentante,                                                                     " & vbCrLf &
                   "   			Clientes.Cliente_Id AS CnpjCliente,                                                                          " & vbCrLf &
                   "   			Clientes.Nome AS NomeDoCliente,                                                                              " & vbCrLf &
                   "   			Clientes.Cidade AS CidadeDoCliente,                                                                          " & vbCrLf &
                   "   			Comissoes.Pedido_Id AS Pedido,                                                                               " & vbCrLf &
                   "   			Pedidos.DataPedido,                                                                                          " & vbCrLf &
                   "   			SUM(NotasFiscaisXItens.QuantidadeFiscal) AS QuantidadeFaturada,                                              " & vbCrLf &
                   "   			SUM(NotasFiscaisXItens.Valor) AS ValorFaturado,                                                              " & vbCrLf &
                   "   			Comissoes.Percentual,                                                                                        " & vbCrLf &
                   "   			Case                                                                                                         " & vbCrLf &
                   "   			when Pedidos.Operacao <> 36                                                                                  " & vbCrLf &
                   "   				then SUM(NotasFiscaisXItens.Valor) * Comissoes.Percentual / 100                                          " & vbCrLf &
                   "   				else ((SUM(NotasFiscaisXItens.QuantidadeFiscal / 1000) * Comissoes.Percentual) * Cotacoes.Indice)        " & vbCrLf &
                   "   						* (((case                                                                                        " & vbCrLf &
                   "   							 when M.Classificacao = 'O'                                                                  " & vbCrLf &
                   "   									then VLDOCUMENTO.Valoroficial                                                        " & vbCrLf &
                   "   									else VLDOCUMENTO.ValorMoeda                                                          " & vbCrLf &
                   "   							 end --GET VALOR DOCUMENTO                                                                   " & vbCrLf &
                   "   						/ SUM(NotasFiscaisXItens.Valor) * 100)) / 100)                                                   " & vbCrLf &
                   "   			End  AS Comissao,                                                                                            " & vbCrLf &
                   "   			CP.Titulo_Id AS Titulo,                                                                                      " & vbCrLf &
                   "   			CP.DataBaixa AS DataTitulo,                                                                                  " & vbCrLf &
                   "   			CASE WHEN CP.Provisao = 1                                                                                    " & vbCrLf &
                   "   				THEN case                                                                                                " & vbCrLf &
                   "   						when M.Classificacao = 'O'                                                                       " & vbCrLf &
                   "   							then VLDOCUMENTO.Valoroficial                                                                " & vbCrLf &
                   "   							else VLDOCUMENTO.ValorMoeda                                                                  " & vbCrLf &
                   "   					 end --GET VALOR DOCUMENTO                                                                           " & vbCrLf &
                   "   				ELSE 0                                                                                                   " & vbCrLf &
                   "   			END AS ValorPago,                                                                                            " & vbCrLf &
                   "   			CASE WHEN CP.Provisao = 1                                                                                    " & vbCrLf &
                   "   				THEN case                                                                                                " & vbCrLf &
                   "   						when M.Classificacao = 'O'                                                                       " & vbCrLf &
                   "   						then VLDOCUMENTO.Valoroficial                                                                    " & vbCrLf &
                   "   						else VLDOCUMENTO.ValorMoeda                                                                      " & vbCrLf &
                   "   					 end --GET VALOR DOCUMENTO                                                                           " & vbCrLf &
                   "   				ELSE 0                                                                                                   " & vbCrLf &
                   "   			END * Comissoes.Percentual / 100 AS ValorLiberado,                                                           " & vbCrLf &
                   "                       Pedidos.Operacao,                                                                                 " & vbCrLf &
                   "                       Cotacoes.Indice,                                                                                  " & vbCrLf &
                   "                      CP.RecPag as Conta                                                                                 " & vbCrLf &
                   "                       FROM Comissoes                                                                                    " & vbCrLf &
                   "   	   LEFT JOIN Titulos CP                                                                                              " & vbCrLf &
                   "   	     ON Comissoes.Empresa_Id    = CP.Empresa                                                                         " & vbCrLf &
                   "           AND Comissoes.EndEmpresa_Id = CP.EndEmpresa                                                                   " & vbCrLf &
                   "           AND Comissoes.Pedido_Id     = CP.Pedido                                                                       " & vbCrLf &
                   "   	  INNER JOIN TitulosxContaContabil VLDOCUMENTO                                                                       " & vbCrLf &
                   "   		 on CP.Titulo_Id           = VLDOCUMENTO.Titulo_Id                                                               " & vbCrLf &
                   "   		and CP.ContaContabilCliFor = VLDOCUMENTO.Conta_Id                                                                " & vbCrLf &
                   "   		and VLDOCUMENTO.DC_Id      = case                                                                                " & vbCrLf &
                   "   									   when CP.RecPag in ('P','C')                                                       " & vbCrLf &
                   "   									     then 'D'                                                                        " & vbCrLf &
                   "   										 else 'C'                                                                        " & vbCrLf &
                   "                       End                                                                                               " & vbCrLf &
                   "   	  INNER JOIN Cotacoes                                                                                                " & vbCrLf &
                   "   		 ON CP.DataBaixa = Cotacoes.Data_Id                                                                              " & vbCrLf &
                   "           AND CP.Indexador = Cotacoes.Indexador_Id                                                                      " & vbCrLf &
                   "         INNER JOIN Pedidos                                                                                              " & vbCrLf &
                   "            ON Comissoes.Pedido_Id     = Pedidos.Pedido_Id                                                               " & vbCrLf &
                   "           AND Comissoes.Empresa_Id    = Pedidos.Empresa_Id                                                              " & vbCrLf &
                   "           AND Comissoes.EndEmpresa_Id = Pedidos.EndEmpresa_Id                                                           " & vbCrLf &
                   "          LEFT JOIN Clientes                                                                                             " & vbCrLf &
                   "            ON Clientes.Cliente_Id  = Pedidos.Cliente                                                                    " & vbCrLf &
                   "           AND Clientes.Endereco_Id = Pedidos.EndCliente                                                                 " & vbCrLf &
                   "         INNER JOIN NotasFiscaisXItens                                                                                   " & vbCrLf &
                   "            ON NotasFiscaisXItens.Pedido = Pedidos.Pedido_Id                                                             " & vbCrLf &
                   "          LEFT JOIN Clientes AS Representante                                                                            " & vbCrLf &
                   "            ON Comissoes.Representante_Id    = Representante.Cliente_Id                                                  " & vbCrLf &
                   "           AND Comissoes.EndRepresentante_Id = Representante.Endereco_Id                                                 " & vbCrLf &
                   "          LEFT JOIN Clientes AS Empresa                                                                                  " & vbCrLf &
                   "            ON Comissoes.Empresa_Id    = Empresa.Cliente_Id                                                              " & vbCrLf &
                   "           AND Comissoes.EndEmpresa_Id = Empresa.Endereco_Id                                                             " & vbCrLf &
                   "                                                                                                                         " & vbCrLf &
                   "           INNER JOIN Moedas M                                                                                           " & vbCrLf &
                   "   			on M.Moeda_id = CP.Moeda                                                                                     " & vbCrLf &
                   "                                                                                                                         " & vbCrLf &
                   "                       WHERE(Pedidos.Situacao = 1)                                                                       " & vbCrLf &
                   "       and isnull(CP.RegistroMestre,0) <> CP.Titulo_Id                                                                   " & vbCrLf &
                   "       And CP.Baixa  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf &
                   "       and CP.RecPag = 'P'                                                                                                  " & vbCrLf &
                   "    and Comissoes.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf &
                   "    and Comissoes.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

            If ddlRepresentante.SelectedIndex > 0 Then
                Sql &= "And Comissoes.Representante_Id = '" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            Sql &= "    GROUP BY Empresa.Reduzido,                                             " & vbCrLf &
            "   			Comissoes.Empresa_Id ,                                             " & vbCrLf &
            "   			Comissoes.EndEmpresa_Id ,                                          " & vbCrLf &
            "   			Empresa.Nome ,                                                     " & vbCrLf &
            "   			Empresa.Cidade ,                                                   " & vbCrLf &
            "   			Empresa.Estado ,                                                   " & vbCrLf &
            "   			Comissoes.Representante_Id ,                                       " & vbCrLf &
            "   			Comissoes.EndRepresentante_Id ,                                    " & vbCrLf &
            "   			Representante.Nome,                                                " & vbCrLf &
            "   			Representante.Cidade,                                              " & vbCrLf &
            "   			Representante.Estado,                                              " & vbCrLf &
            "   			Comissoes.Pedido_Id,                                               " & vbCrLf &
            "   			Comissoes.Percentual,                                              " & vbCrLf &
            "   			Pedidos.DataPedido,                                                " & vbCrLf &
            "   			CP.Titulo_Id,                                                      " & vbCrLf &
            "               CP.RecPag,                                                         " & vbCrLf &
            "   			CP.DataBaixa,			                                           " & vbCrLf &
            "   			Clientes.Cliente_Id,                                               " & vbCrLf &
            "   			Clientes.Nome,                                                     " & vbCrLf &
            "   			Clientes.Cidade,                                                   " & vbCrLf &
            "   			CP.Provisao,			                                           " & vbCrLf &
            "   			Pedidos.Operacao,                                                  " & vbCrLf &
            "   			Cotacoes.Indice,                                                   " & vbCrLf &
            "   			M.Classificacao,                                                   " & vbCrLf &
            "   			VLDOCUMENTO.Valoroficial,                                          " & vbCrLf &
            "   			VLDOCUMENTO.ValorMoeda                                             " & vbCrLf &
            "                                                                                  " & vbCrLf &
            "     Having   ISNULL(SUM(CASE                                                     " & vbCrLf &
            "   						WHEN CP.Provisao in(1,2)                               " & vbCrLf &
            "   							THEN case                                          " & vbCrLf &
            "   									when M.Classificacao = 'O'                 " & vbCrLf &
            "   										then VLDOCUMENTO.Valoroficial          " & vbCrLf &
            "           Else VLDOCUMENTO.ValorMoeda                                            " & vbCrLf &
            "               End                                                                " & vbCrLf &
            "   							ELSE 0                                             " & vbCrLf &
            "   						END), 0) <> 0                                          " & vbCrLf

        Else
            Sql = " SELECT Empresa.Reduzido AS RedEmpresa, Comissoes.Empresa_Id AS CNPJEmpresa, Comissoes.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa," & vbCrLf &
                "           Empresa.Cidade AS CidadeEmpresa, Empresa.Estado AS UFEmpresa, Comissoes.Representante_Id AS CNPJRepresentante," & vbCrLf &
                "           Comissoes.EndRepresentante_Id AS EndRepresentante, Representante.Nome AS NomeRepresentante, Representante.Cidade AS CidadeRepresentante," & vbCrLf &
                "           Representante.Estado AS UFRepresentante, Clientes.Cliente_Id AS CnpjCliente, Clientes.Nome AS NomeDoCliente, Clientes.Cidade AS CidadeDoCliente," & vbCrLf &
                "           Comissoes.Pedido_Id AS Pedido, Pedidos.DataPedido, NFI.Nota_id AS Nota," & vbCrLf &
                "           SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.QuantidadeFiscal ELSE NFI.QuantidadeFiscal * - 1 END)" & vbCrLf &
                "           AS QuantidadeFaturada, SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * - 1 END)" & vbCrLf &
                "           AS ValorFaturado,                                                                                           " & vbCrLf &
                "           SUM(ISNULL(EncFunrural.Valor,0)) AS Funrural,                                                               " & vbCrLf &
                "           SUM(ISNULL(EncSenar.Valor,0)) AS Senar,                                                                     " & vbCrLf &
                "           SUM(ISNULL(EncIcms.Base,0)) AS BaseIcms,                                                                    " & vbCrLf &
                "           SUM(ISNULL(EncIcms.Valor,0)) AS Icms,                                                                       " & vbCrLf &
                "           SUM(ISNULL(EncIcmsST.Base,0)) AS BaseIcmsST,                                                                " & vbCrLf &
                "           SUM(ISNULL(EncIcmsST.Valor,0)) AS IcmsST,                                                                   " & vbCrLf &
                "           SUM(ISNULL(EncPis.Base,0)) AS BasePis,                                                                      " & vbCrLf &
                "           SUM(ISNULL(EncPis.Valor,0)) AS Pis,                                                                         " & vbCrLf &
                "           SUM(ISNULL(EncCofins.Base,0)) AS BaseCofins,                                                                " & vbCrLf &
                "           SUM(ISNULL(EncCofins.Valor,0)) AS Cofins,                                                                   " & vbCrLf &
                "           SUM(ISNULL(EncIpi.Base,0)) AS BaseIpi,                                                                      " & vbCrLf &
                "           SUM(ISNULL(EncIpi.Valor,0)) AS Ipi,                                                                         " & vbCrLf &
                "           SUM(ISNULL(EncDesconto.Valor,0)) AS Desconto,                                                               " & vbCrLf &
                "           Comissoes.Percentual,                                                                                       " & vbCrLf &
                "           CASE WHEN Pedidos.Operacao <> 36 THEN SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * - 1 END)" & vbCrLf &
                "            * Comissoes.Percentual / 100 ELSE ((SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.QuantidadeFiscal ELSE NFI.QuantidadeFiscal" & vbCrLf &
                "            * - 1 END / 1000) * Comissoes.Percentual) * isnull(Cotacoes.Indice, 0)) * (((ContasAReceber.ValorDoDocumento / SUM(NFI.Valor) * 100)) / 100)" & vbCrLf &
                "           END AS Comissao, ContasAReceber.Registro_Id AS Titulo, ContasAReceber.Baixa AS DataTitulo, sum(ContasAReceber.ValorDoDocumento) AS ValorTitulo," & vbCrLf &
                "           CASE WHEN ContasAReceber.Provisao = 1 THEN ValorDoDocumento ELSE 0 END AS ValorPago," & vbCrLf &
                "           CASE WHEN ContasAReceber.Provisao = 1 THEN ValorDoDocumento ELSE 0 END * Comissoes.Percentual / 100 AS ValorLiberado,   " & vbCrLf &
                "           SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * -1 END)                                      " & vbCrLf &
                "               -   (                                                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncFunrural.Valor,0)) +                                                                          " & vbCrLf &
                "                       SUM(ISNULL(EncSenar.Valor,0)) +                                                                             " & vbCrLf &
                "                       SUM(ISNULL(EncIcms.Valor,0)) +                                                                              " & vbCrLf &
                "                       SUM(ISNULL(EncIcmsST.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncPis.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncCofins.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncIpi.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncDesconto.Valor,0))                                                                            " & vbCrLf &
                "                   )                                                                                                               " & vbCrLf &
                "           AS ValorLiquido,                                                                                                        " & vbCrLf &
                "           (SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * -1 END)                                     " & vbCrLf &
                "               -   (                                                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncFunrural.Valor,0)) +                                                                          " & vbCrLf &
                "                       SUM(ISNULL(EncSenar.Valor,0)) +                                                                             " & vbCrLf &
                "                       SUM(ISNULL(EncIcms.Valor,0)) +                                                                              " & vbCrLf &
                "                       SUM(ISNULL(EncIcmsST.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncPis.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncCofins.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncIpi.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncDesconto.Valor,0))                                                                            " & vbCrLf &
                "                   )) * Comissoes.Percentual / 100                                                                                 " & vbCrLf &
                "           AS ComissaoLiquida,                                                                                                     " & vbCrLf &
                " Pedidos.Operacao,                                                                                                                 " & vbCrLf &
                "           isnull(Cotacoes.Indice, 0) as Indice, 'R' as Conta" & vbCrLf &
                " FROM  SubOperacoes INNER JOIN" & vbCrLf &
                "            Clientes INNER JOIN" & vbCrLf &
                "            Pedidos ON Clientes.Cliente_Id = Pedidos.Cliente AND Clientes.Endereco_Id = Pedidos.EndCliente INNER JOIN" & vbCrLf &
                "            NotasFiscaisXItens AS NFI ON Pedidos.Pedido_Id = NFI.Pedido And Pedidos.Empresa_ID = NFI.Empresa_Id " & vbCrLf &
                "            And Pedidos.EndEmpresa_ID = NFI.EndEmpresa_Id " & vbCrLf &
                "           INNER JOIN OperacaoXEstadoXEncargo oeP                                                                      " & vbCrLf &
                "                   ON oeP.Codigo_Id = NFI.OperacaoXEstado                                                              " & vbCrLf &
                "                  AND oeP.Encargo_Id = 'PRODUTO'                                                                       " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Valor                                                   " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'FUNRURAL'                                                                    " & vbCrLf &
                "                        AND Valor > 0 ) as EncFunrural                                                                 " & vbCrLf &
                "                   ON EncFunrural.Empresa_Id = NFI.Empresa_Id                                                          " & vbCrLf &
                "                  AND EncFunrural.EndEmpresa_Id = NFI.EndEmpresa_Id                                                    " & vbCrLf &
                "                  AND EncFunrural.Cliente_Id = NFI.Cliente_Id                                                          " & vbCrLf &
                "                  AND EncFunrural.EndCliente_Id = NFI.EndCliente_Id                                                    " & vbCrLf &
                "                  AND EncFunrural.EntradaSaida_Id = NFI.EntradaSaida_Id                                                " & vbCrLf &
                "                  AND EncFunrural.Serie_Id = NFI.Serie_Id                                                              " & vbCrLf &
                "                  AND EncFunrural.Nota_Id = NFI.Nota_Id                                                                " & vbCrLf &
                "                  AND EncFunrural.Produto_Id = NFI.Produto_Id                                                          " & vbCrLf &
                "                  AND EncFunrural.CFOP_Id = NFI.CFOP_Id                                                                " & vbCrLf &
                "                  AND EncFunrural.Sequencia_Id = NFI.Sequencia_Id                                                      " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Valor                                                   " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'SENAR'                                                                      " & vbCrLf &
                "                        AND Valor > 0 ) as EncSenar                                                                    " & vbCrLf &
                "                   ON EncSenar.Empresa_Id = NFI.Empresa_Id                                                             " & vbCrLf &
                "                  AND EncSenar.EndEmpresa_Id = NFI.EndEmpresa_Id                                                       " & vbCrLf &
                "                  AND EncSenar.Cliente_Id = NFI.Cliente_Id                                                             " & vbCrLf &
                "                  AND EncSenar.EndCliente_Id = NFI.EndCliente_Id                                                       " & vbCrLf &
                "                  AND EncSenar.EntradaSaida_Id = NFI.EntradaSaida_Id                                                   " & vbCrLf &
                "                  AND EncSenar.Serie_Id = NFI.Serie_Id                                                                 " & vbCrLf &
                "                  AND EncSenar.Nota_Id = NFI.Nota_Id                                                                   " & vbCrLf &
                "                  AND EncSenar.Produto_Id = NFI.Produto_Id                                                             " & vbCrLf &
                "                  AND EncSenar.CFOP_Id = NFI.CFOP_Id                                                                   " & vbCrLf &
                "                  AND EncSenar.Sequencia_Id = NFI.Sequencia_Id                                                         " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'ICMS'                                                                       " & vbCrLf &
                "                        AND Valor > 0 ) as EncIcms                                                                     " & vbCrLf &
                "                   ON EncIcms.Empresa_Id = NFI.Empresa_Id                                                              " & vbCrLf &
                "                  AND EncIcms.EndEmpresa_Id = NFI.EndEmpresa_Id                                                        " & vbCrLf &
                "                  AND EncIcms.Cliente_Id = NFI.Cliente_Id                                                              " & vbCrLf &
                "                  AND EncIcms.EndCliente_Id = NFI.EndCliente_Id                                                        " & vbCrLf &
                "                  AND EncIcms.EntradaSaida_Id = NFI.EntradaSaida_Id                                                    " & vbCrLf &
                "                  AND EncIcms.Serie_Id = NFI.Serie_Id                                                                  " & vbCrLf &
                "                  AND EncIcms.Nota_Id = NFI.Nota_Id                                                                    " & vbCrLf &
                "                  AND EncIcms.Produto_Id = NFI.Produto_Id                                                              " & vbCrLf &
                "                  AND EncIcms.CFOP_Id = NFI.CFOP_Id                                                                    " & vbCrLf &
                "                  AND EncIcms.Sequencia_Id = NFI.Sequencia_Id                                                          " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'ICMS-ST'                                                                    " & vbCrLf &
                "                        AND Valor > 0 ) as EncIcmsST                                                                   " & vbCrLf &
                "                   ON EncIcmsST.Empresa_Id = NFI.Empresa_Id                                                            " & vbCrLf &
                "                  AND EncIcmsST.EndEmpresa_Id = NFI.EndEmpresa_Id                                                      " & vbCrLf &
                "                  AND EncIcmsST.Cliente_Id = NFI.Cliente_Id                                                            " & vbCrLf &
                "                  AND EncIcmsST.EndCliente_Id = NFI.EndCliente_Id                                                      " & vbCrLf &
                "                  AND EncIcmsST.EntradaSaida_Id = NFI.EntradaSaida_Id                                                  " & vbCrLf &
                "                  AND EncIcmsST.Serie_Id = NFI.Serie_Id                                                                " & vbCrLf &
                "                  AND EncIcmsST.Nota_Id = NFI.Nota_Id                                                                  " & vbCrLf &
                "                  AND EncIcmsST.Produto_Id = NFI.Produto_Id                                                            " & vbCrLf &
                "                  AND EncIcmsST.CFOP_Id = NFI.CFOP_Id                                                                  " & vbCrLf &
                "                  AND EncIcmsST.Sequencia_Id = NFI.Sequencia_Id                                                        " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'PIS'                                                                        " & vbCrLf &
                "                        AND Valor > 0 ) as EncPis                                                                      " & vbCrLf &
                "                   ON EncPis.Empresa_Id = NFI.Empresa_Id                                                               " & vbCrLf &
                "                  AND EncPis.EndEmpresa_Id = NFI.EndEmpresa_Id                                                         " & vbCrLf &
                "                  AND EncPis.Cliente_Id = NFI.Cliente_Id                                                               " & vbCrLf &
                "                  AND EncPis.EndCliente_Id = NFI.EndCliente_Id                                                         " & vbCrLf &
                "                  AND EncPis.EntradaSaida_Id = NFI.EntradaSaida_Id                                                     " & vbCrLf &
                "                  AND EncPis.Serie_Id = NFI.Serie_Id                                                                   " & vbCrLf &
                "                  AND EncPis.Nota_Id = NFI.Nota_Id                                                                     " & vbCrLf &
                "                  AND EncPis.Produto_Id = NFI.Produto_Id                                                               " & vbCrLf &
                "                  AND EncPis.CFOP_Id = NFI.CFOP_Id                                                                     " & vbCrLf &
                "                  AND EncPis.Sequencia_Id = NFI.Sequencia_Id                                                           " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'COFINS'                                                                     " & vbCrLf &
                "                        AND Valor > 0 ) as EncCofins                                                                   " & vbCrLf &
                "                   ON EncCofins.Empresa_Id = NFI.Empresa_Id                                                             " & vbCrLf &
                "                  AND EncCofins.EndEmpresa_Id = NFI.EndEmpresa_Id                                                       " & vbCrLf &
                "                  AND EncCofins.Cliente_Id = NFI.Cliente_Id                                                             " & vbCrLf &
                "                  AND EncCofins.EndCliente_Id = NFI.EndCliente_Id                                                       " & vbCrLf &
                "                  AND EncCofins.EntradaSaida_Id = NFI.EntradaSaida_Id                                                   " & vbCrLf &
                "                  AND EncCofins.Serie_Id = NFI.Serie_Id                                                                 " & vbCrLf &
                "                  AND EncCofins.Nota_Id = NFI.Nota_Id                                                                   " & vbCrLf &
                "                  AND EncCofins.Produto_Id = NFI.Produto_Id                                                             " & vbCrLf &
                "                  AND EncCofins.CFOP_Id = NFI.CFOP_Id                                                                   " & vbCrLf &
                "                  AND EncCofins.Sequencia_Id = NFI.Sequencia_Id                                                         " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                           " & vbCrLf &
                "                      Where Encargo_Id = 'IPI'                                                                         " & vbCrLf &
                "                        AND Valor > 0 ) as EncIpi                                                                      " & vbCrLf &
                "                   ON EncIpi.Empresa_Id = NFI.Empresa_Id                                                               " & vbCrLf &
                "                  AND EncIpi.EndEmpresa_Id = NFI.EndEmpresa_Id                                                         " & vbCrLf &
                "                  AND EncIpi.Cliente_Id = NFI.Cliente_Id                                                               " & vbCrLf &
                "                  AND EncIpi.EndCliente_Id = NFI.EndCliente_Id                                                         " & vbCrLf &
                "                  AND EncIpi.EntradaSaida_Id = NFI.EntradaSaida_Id                                                     " & vbCrLf &
                "                  AND EncIpi.Serie_Id = NFI.Serie_Id                                                                   " & vbCrLf &
                "                  AND EncIpi.Nota_Id = NFI.Nota_Id                                                                     " & vbCrLf &
                "                  AND EncIpi.Produto_Id = NFI.Produto_Id                                                               " & vbCrLf &
                "                  AND EncIpi.CFOP_Id = NFI.CFOP_Id                                                                     " & vbCrLf &
                "                  AND EncIpi.Sequencia_Id = NFI.Sequencia_Id                                                           " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Valor                                                  " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'DESCONTOS'                                                                   " & vbCrLf &
                "                        AND Valor > 0 ) as EncDesconto                                                                 " & vbCrLf &
                "                   ON EncDesconto.Empresa_Id = NFI.Empresa_Id                                                           " & vbCrLf &
                "                  AND EncDesconto.EndEmpresa_Id = NFI.EndEmpresa_Id                                                     " & vbCrLf &
                "                  AND EncDesconto.Cliente_Id = NFI.Cliente_Id                                                           " & vbCrLf &
                "                  AND EncDesconto.EndCliente_Id = NFI.EndCliente_Id                                                     " & vbCrLf &
                "                  AND EncDesconto.EntradaSaida_Id = NFI.EntradaSaida_Id                                                 " & vbCrLf &
                "                  AND EncDesconto.Serie_Id = NFI.Serie_Id                                                               " & vbCrLf &
                "                  AND EncDesconto.Nota_Id = NFI.Nota_Id                                                                 " & vbCrLf &
                "                  AND EncDesconto.Produto_Id = NFI.Produto_Id                                                           " & vbCrLf &
                "                  AND EncDesconto.CFOP_Id = NFI.CFOP_Id                                                                 " & vbCrLf &
                "                  AND EncDesconto.Sequencia_Id = NFI.Sequencia_Id                                                       " & vbCrLf &
                " INNER JOIN" & vbCrLf &
                "            Produtos ON NFI.Produto_Id = Produtos.Produto_Id ON SubOperacoes.Operacao_Id = NFI.Operacao And" & vbCrLf &
                "            SubOperacoes.SubOperacoes_Id = NFI.SubOperacao RIGHT OUTER JOIN" & vbCrLf &
                "            Comissoes LEFT OUTER JOIN" & vbCrLf &
                "            ContasAReceber Left JOIN" & vbCrLf &
                "            Cotacoes ON ContasAReceber.Baixa = Cotacoes.Data_Id And ContasAReceber.Indexador = Cotacoes.Indexador_Id ON" & vbCrLf &
                "            Comissoes.Empresa_Id = ContasAReceber.EmpresaPedido And Comissoes.EndEmpresa_Id = ContasAReceber.EndEmpresaPedido And" & vbCrLf &
                "            Comissoes.Pedido_Id = ContasAReceber.Pedido ON Pedidos.Pedido_Id = Comissoes.Pedido_Id And Pedidos.Empresa_Id = Comissoes.Empresa_Id And" & vbCrLf &
                "            Pedidos.EndEmpresa_Id = Comissoes.EndEmpresa_Id LEFT OUTER JOIN" & vbCrLf &
                "            Clientes AS Representante ON Comissoes.Representante_Id = Representante.Cliente_Id And" & vbCrLf &
                "            Comissoes.EndRepresentante_Id = Representante.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
                "             Clientes AS Empresa ON Comissoes.Empresa_Id = Empresa.Cliente_Id And Comissoes.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf &
                " WHERE     (Pedidos.Situacao = 1) And  ContasAReceber.Grupado <> 'M'" & vbCrLf &
                "           And (ContasAReceber.Baixa  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

            If ddlRepresentante.SelectedIndex > 0 Then
                Sql &= "        And Comissoes.Representante_Id = '" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            Sql &= " GROUP BY Comissoes.Empresa_Id, Comissoes.EndEmpresa_Id, Empresa.Reduzido, Empresa.Nome, Empresa.Cidade, Empresa.Estado, Comissoes.Representante_Id," & vbCrLf &
                "            Comissoes.EndRepresentante_Id, Representante.Nome, Representante.Cidade, Representante.Estado, Comissoes.Pedido_Id, Comissoes.Percentual," & vbCrLf &
                "            Comissoes.ValorComissao, Comissoes.Principal, Pedidos.DataPedido, ContasAReceber.Registro_Id, ContasAReceber.Baixa, Clientes.Cliente_Id, Clientes.Nome," & vbCrLf &
                "            Clientes.Cidade, ContasAReceber.Provisao, ContasAReceber.ValorDoDocumento, Cotacoes.Indexador_Id, Cotacoes.Indice, Pedidos.Operacao," & vbCrLf &
                "            ContasAReceber.Carteira, NFI.Nota_id" & vbCrLf &
                " HAVING (ISNULL(SUM(CASE WHEN ContasAReceber.Provisao IN (1,2) THEN ValorDoDocumento ELSE 0 END), 0) <> 0)" & vbCrLf &
                "" & vbCrLf &
                "  Union ALL" & vbCrLf &
                "" & vbCrLf &
                " SELECT Empresa.Reduzido AS RedEmpresa, Comissoes.Empresa_Id AS CNPJEmpresa, Comissoes.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa," & vbCrLf &
                "            Empresa.Cidade AS CidadeEmpresa, Empresa.Estado AS UFEmpresa, Comissoes.Representante_Id AS CNPJRepresentante," & vbCrLf &
                "            Comissoes.EndRepresentante_Id AS EndRepresentante, Representante.Nome AS NomeRepresentante, Representante.Cidade AS CidadeRepresentante," & vbCrLf &
                "            Representante.Estado AS UFRepresentante, Clientes.Cliente_Id AS CnpjCliente, Clientes.Nome AS NomeDoCliente, Clientes.Cidade AS CidadeDoCliente," & vbCrLf &
                "            Comissoes.Pedido_Id AS Pedido, Pedidos.DataPedido, NFI.Nota_id AS Nota," & vbCrLf &
                "            SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.QuantidadeFiscal ELSE NFI.QuantidadeFiscal * - 1 END)" & vbCrLf &
                "            AS QuantidadeFaturada, SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * - 1 END)" & vbCrLf &
                "           AS ValorFaturado,                                                                                           " & vbCrLf &
                "           SUM(ISNULL(EncFunrural.Valor,0)) AS Funrural,                                                               " & vbCrLf &
                "           SUM(ISNULL(EncSenar.Valor,0)) AS Senar,                                                                     " & vbCrLf &
                "           SUM(ISNULL(EncIcms.Base,0)) AS BaseIcms,                                                                    " & vbCrLf &
                "           SUM(ISNULL(EncIcms.Valor,0)) AS Icms,                                                                       " & vbCrLf &
                "           SUM(ISNULL(EncIcmsST.Base,0)) AS BaseIcmsST,                                                                " & vbCrLf &
                "           SUM(ISNULL(EncIcmsST.Valor,0)) AS IcmsST,                                                                   " & vbCrLf &
                "           SUM(ISNULL(EncPis.Base,0)) AS BasePis,                                                                      " & vbCrLf &
                "           SUM(ISNULL(EncPis.Valor,0)) AS Pis,                                                                         " & vbCrLf &
                "           SUM(ISNULL(EncCofins.Base,0)) AS BaseCofins,                                                                " & vbCrLf &
                "           SUM(ISNULL(EncCofins.Valor,0)) AS Cofins,                                                                   " & vbCrLf &
                "           SUM(ISNULL(EncIpi.Base,0)) AS BaseIpi,                                                                      " & vbCrLf &
                "           SUM(ISNULL(EncIpi.Valor,0)) AS Ipi,                                                                         " & vbCrLf &
                "           SUM(ISNULL(EncDesconto.Valor,0)) AS Desconto,                                                               " & vbCrLf &
                "           Comissoes.Percentual,                                                                                       " & vbCrLf &
                "            CASE WHEN Pedidos.Operacao <> 36 THEN SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * - 1 END)" & vbCrLf &
                "             * Comissoes.Percentual / 100 ELSE ((SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.QuantidadeFiscal ELSE NFI.QuantidadeFiscal" & vbCrLf &
                "             * - 1 END / 1000) * Comissoes.Percentual) * isnull(Cotacoes.Indice, 0))" & vbCrLf &
                "            * (((ContasAPagar.ValorDoDocumento / SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * - 1 END)" & vbCrLf &
                "             * 100)) / 100) END AS Comissao, ContasAPagar.Registro_Id AS Titulo, ContasAPagar.Baixa AS DataTitulo, sum(ContasAPagar.ValorDoDocumento) AS ValorTitulo," & vbCrLf &
                "            CASE WHEN ContasAPagar.Provisao = 1 THEN ValorDoDocumento ELSE 0 END AS ValorPago," & vbCrLf &
                "            CASE WHEN ContasAPagar.Provisao = 1 THEN ValorDoDocumento ELSE 0 END * Comissoes.Percentual / 100 AS ValorLiberado,    " & vbCrLf &
                "           SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * -1 END)                                      " & vbCrLf &
                "               -   (                                                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncFunrural.Valor,0)) +                                                                          " & vbCrLf &
                "                       SUM(ISNULL(EncSenar.Valor,0)) +                                                                             " & vbCrLf &
                "                       SUM(ISNULL(EncIcms.Valor,0)) +                                                                              " & vbCrLf &
                "                       SUM(ISNULL(EncIcmsST.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncPis.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncCofins.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncIpi.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncDesconto.Valor,0))                                                                            " & vbCrLf &
                "                   )                                                                                                               " & vbCrLf &
                "           AS ValorLiquido,                                                                                                        " & vbCrLf &
                "           (SUM(CASE WHEN SubOperacoes.Devolucao = 'N' THEN NFI.Valor ELSE NFI.Valor * -1 END)                                     " & vbCrLf &
                "               -   (                                                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncFunrural.Valor,0)) +                                                                          " & vbCrLf &
                "                       SUM(ISNULL(EncSenar.Valor,0)) +                                                                             " & vbCrLf &
                "                       SUM(ISNULL(EncIcms.Valor,0)) +                                                                              " & vbCrLf &
                "                       SUM(ISNULL(EncIcmsST.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncPis.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncCofins.Valor,0)) +                                                                            " & vbCrLf &
                "                       SUM(ISNULL(EncIpi.Valor,0)) +                                                                               " & vbCrLf &
                "                       SUM(ISNULL(EncDesconto.Valor,0))                                                                            " & vbCrLf &
                "                   )) * Comissoes.Percentual / 100                                                                                 " & vbCrLf &
                "           AS ComissaoLiquida,                                                                                                     " & vbCrLf &
                " Pedidos.Operacao,                                                                                                                 " & vbCrLf &
                "            isnull(Cotacoes.Indice, 0) as Indice, 'P' as Conta" & vbCrLf &
                " FROM  SubOperacoes INNER JOIN" & vbCrLf &
                "            Clientes INNER JOIN" & vbCrLf &
                "            Pedidos ON Clientes.Cliente_Id = Pedidos.Cliente AND Clientes.Endereco_Id = Pedidos.EndCliente INNER JOIN" & vbCrLf &
                "            NotasFiscaisXItens AS NFI ON Pedidos.Pedido_Id = NFI.Pedido And Pedidos.Empresa_ID = NFI.Empresa_Id " & vbCrLf &
                "            And Pedidos.EndEmpresa_ID = NFI.EndEmpresa_Id " & vbCrLf &
                "           INNER JOIN OperacaoXEstadoXEncargo oeP                                                                      " & vbCrLf &
                "                   ON oeP.Codigo_Id = NFI.OperacaoXEstado                                                              " & vbCrLf &
                "                  AND oeP.Encargo_Id = 'PRODUTO'                                                                       " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Valor                                                   " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'FUNRURAL'                                                                    " & vbCrLf &
                "                        AND Valor > 0 ) as EncFunrural                                                                 " & vbCrLf &
                "                   ON EncFunrural.Empresa_Id = NFI.Empresa_Id                                                          " & vbCrLf &
                "                  AND EncFunrural.EndEmpresa_Id = NFI.EndEmpresa_Id                                                    " & vbCrLf &
                "                  AND EncFunrural.Cliente_Id = NFI.Cliente_Id                                                          " & vbCrLf &
                "                  AND EncFunrural.EndCliente_Id = NFI.EndCliente_Id                                                    " & vbCrLf &
                "                  AND EncFunrural.EntradaSaida_Id = NFI.EntradaSaida_Id                                                " & vbCrLf &
                "                  AND EncFunrural.Serie_Id = NFI.Serie_Id                                                              " & vbCrLf &
                "                  AND EncFunrural.Nota_Id = NFI.Nota_Id                                                                " & vbCrLf &
                "                  AND EncFunrural.Produto_Id = NFI.Produto_Id                                                          " & vbCrLf &
                "                  AND EncFunrural.CFOP_Id = NFI.CFOP_Id                                                                " & vbCrLf &
                "                  AND EncFunrural.Sequencia_Id = NFI.Sequencia_Id                                                      " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Valor                                                   " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'SENAR'                                                                      " & vbCrLf &
                "                        AND Valor > 0 ) as EncSenar                                                                    " & vbCrLf &
                "                   ON EncSenar.Empresa_Id = NFI.Empresa_Id                                                             " & vbCrLf &
                "                  AND EncSenar.EndEmpresa_Id = NFI.EndEmpresa_Id                                                       " & vbCrLf &
                "                  AND EncSenar.Cliente_Id = NFI.Cliente_Id                                                             " & vbCrLf &
                "                  AND EncSenar.EndCliente_Id = NFI.EndCliente_Id                                                       " & vbCrLf &
                "                  AND EncSenar.EntradaSaida_Id = NFI.EntradaSaida_Id                                                   " & vbCrLf &
                "                  AND EncSenar.Serie_Id = NFI.Serie_Id                                                                 " & vbCrLf &
                "                  AND EncSenar.Nota_Id = NFI.Nota_Id                                                                   " & vbCrLf &
                "                  AND EncSenar.Produto_Id = NFI.Produto_Id                                                             " & vbCrLf &
                "                  AND EncSenar.CFOP_Id = NFI.CFOP_Id                                                                   " & vbCrLf &
                "                  AND EncSenar.Sequencia_Id = NFI.Sequencia_Id                                                         " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'ICMS'                                                                       " & vbCrLf &
                "                        AND Valor > 0 ) as EncIcms                                                                     " & vbCrLf &
                "                   ON EncIcms.Empresa_Id = NFI.Empresa_Id                                                              " & vbCrLf &
                "                  AND EncIcms.EndEmpresa_Id = NFI.EndEmpresa_Id                                                        " & vbCrLf &
                "                  AND EncIcms.Cliente_Id = NFI.Cliente_Id                                                              " & vbCrLf &
                "                  AND EncIcms.EndCliente_Id = NFI.EndCliente_Id                                                        " & vbCrLf &
                "                  AND EncIcms.EntradaSaida_Id = NFI.EntradaSaida_Id                                                    " & vbCrLf &
                "                  AND EncIcms.Serie_Id = NFI.Serie_Id                                                                  " & vbCrLf &
                "                  AND EncIcms.Nota_Id = NFI.Nota_Id                                                                    " & vbCrLf &
                "                  AND EncIcms.Produto_Id = NFI.Produto_Id                                                              " & vbCrLf &
                "                  AND EncIcms.CFOP_Id = NFI.CFOP_Id                                                                    " & vbCrLf &
                "                  AND EncIcms.Sequencia_Id = NFI.Sequencia_Id                                                          " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'ICMS-ST'                                                                    " & vbCrLf &
                "                        AND Valor > 0 ) as EncIcmsST                                                                   " & vbCrLf &
                "                   ON EncIcmsST.Empresa_Id = NFI.Empresa_Id                                                            " & vbCrLf &
                "                  AND EncIcmsST.EndEmpresa_Id = NFI.EndEmpresa_Id                                                      " & vbCrLf &
                "                  AND EncIcmsST.Cliente_Id = NFI.Cliente_Id                                                            " & vbCrLf &
                "                  AND EncIcmsST.EndCliente_Id = NFI.EndCliente_Id                                                      " & vbCrLf &
                "                  AND EncIcmsST.EntradaSaida_Id = NFI.EntradaSaida_Id                                                  " & vbCrLf &
                "                  AND EncIcmsST.Serie_Id = NFI.Serie_Id                                                                " & vbCrLf &
                "                  AND EncIcmsST.Nota_Id = NFI.Nota_Id                                                                  " & vbCrLf &
                "                  AND EncIcmsST.Produto_Id = NFI.Produto_Id                                                            " & vbCrLf &
                "                  AND EncIcmsST.CFOP_Id = NFI.CFOP_Id                                                                  " & vbCrLf &
                "                  AND EncIcmsST.Sequencia_Id = NFI.Sequencia_Id                                                        " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'PIS'                                                                        " & vbCrLf &
                "                        AND Valor > 0 ) as EncPis                                                                      " & vbCrLf &
                "                   ON EncPis.Empresa_Id = NFI.Empresa_Id                                                               " & vbCrLf &
                "                  AND EncPis.EndEmpresa_Id = NFI.EndEmpresa_Id                                                         " & vbCrLf &
                "                  AND EncPis.Cliente_Id = NFI.Cliente_Id                                                               " & vbCrLf &
                "                  AND EncPis.EndCliente_Id = NFI.EndCliente_Id                                                         " & vbCrLf &
                "                  AND EncPis.EntradaSaida_Id = NFI.EntradaSaida_Id                                                     " & vbCrLf &
                "                  AND EncPis.Serie_Id = NFI.Serie_Id                                                                   " & vbCrLf &
                "                  AND EncPis.Nota_Id = NFI.Nota_Id                                                                     " & vbCrLf &
                "                  AND EncPis.Produto_Id = NFI.Produto_Id                                                               " & vbCrLf &
                "                  AND EncPis.CFOP_Id = NFI.CFOP_Id                                                                     " & vbCrLf &
                "                  AND EncPis.Sequencia_Id = NFI.Sequencia_Id                                                           " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'COFINS'                                                                     " & vbCrLf &
                "                        AND Valor > 0 ) as EncCofins                                                                   " & vbCrLf &
                "                   ON EncCofins.Empresa_Id = NFI.Empresa_Id                                                             " & vbCrLf &
                "                  AND EncCofins.EndEmpresa_Id = NFI.EndEmpresa_Id                                                       " & vbCrLf &
                "                  AND EncCofins.Cliente_Id = NFI.Cliente_Id                                                             " & vbCrLf &
                "                  AND EncCofins.EndCliente_Id = NFI.EndCliente_Id                                                       " & vbCrLf &
                "                  AND EncCofins.EntradaSaida_Id = NFI.EntradaSaida_Id                                                   " & vbCrLf &
                "                  AND EncCofins.Serie_Id = NFI.Serie_Id                                                                 " & vbCrLf &
                "                  AND EncCofins.Nota_Id = NFI.Nota_Id                                                                   " & vbCrLf &
                "                  AND EncCofins.Produto_Id = NFI.Produto_Id                                                             " & vbCrLf &
                "                  AND EncCofins.CFOP_Id = NFI.CFOP_Id                                                                   " & vbCrLf &
                "                  AND EncCofins.Sequencia_Id = NFI.Sequencia_Id                                                         " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Base, Valor                                             " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                           " & vbCrLf &
                "                      Where Encargo_Id = 'IPI'                                                                         " & vbCrLf &
                "                        AND Valor > 0 ) as EncIpi                                                                      " & vbCrLf &
                "                   ON EncIpi.Empresa_Id = NFI.Empresa_Id                                                               " & vbCrLf &
                "                  AND EncIpi.EndEmpresa_Id = NFI.EndEmpresa_Id                                                         " & vbCrLf &
                "                  AND EncIpi.Cliente_Id = NFI.Cliente_Id                                                               " & vbCrLf &
                "                  AND EncIpi.EndCliente_Id = NFI.EndCliente_Id                                                         " & vbCrLf &
                "                  AND EncIpi.EntradaSaida_Id = NFI.EntradaSaida_Id                                                     " & vbCrLf &
                "                  AND EncIpi.Serie_Id = NFI.Serie_Id                                                                   " & vbCrLf &
                "                  AND EncIpi.Nota_Id = NFI.Nota_Id                                                                     " & vbCrLf &
                "                  AND EncIpi.Produto_Id = NFI.Produto_Id                                                               " & vbCrLf &
                "                  AND EncIpi.CFOP_Id = NFI.CFOP_Id                                                                     " & vbCrLf &
                "                  AND EncIpi.Sequencia_Id = NFI.Sequencia_Id                                                           " & vbCrLf &
                "           LEFT JOIN (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf &
                "                             Produto_Id, CFOP_Id, Sequencia_Id, Valor                                                  " & vbCrLf &
                "                      from NotasFiscaisXEncargos                                                                       " & vbCrLf &
                "                      Where Encargo_Id = 'DESCONTOS'                                                                   " & vbCrLf &
                "                        AND Valor > 0 ) as EncDesconto                                                                 " & vbCrLf &
                "                   ON EncDesconto.Empresa_Id = NFI.Empresa_Id                                                           " & vbCrLf &
                "                  AND EncDesconto.EndEmpresa_Id = NFI.EndEmpresa_Id                                                     " & vbCrLf &
                "                  AND EncDesconto.Cliente_Id = NFI.Cliente_Id                                                           " & vbCrLf &
                "                  AND EncDesconto.EndCliente_Id = NFI.EndCliente_Id                                                     " & vbCrLf &
                "                  AND EncDesconto.EntradaSaida_Id = NFI.EntradaSaida_Id                                                 " & vbCrLf &
                "                  AND EncDesconto.Serie_Id = NFI.Serie_Id                                                               " & vbCrLf &
                "                  AND EncDesconto.Nota_Id = NFI.Nota_Id                                                                 " & vbCrLf &
                "                  AND EncDesconto.Produto_Id = NFI.Produto_Id                                                           " & vbCrLf &
                "                  AND EncDesconto.CFOP_Id = NFI.CFOP_Id                                                                 " & vbCrLf &
                "                  AND EncDesconto.Sequencia_Id = NFI.Sequencia_Id                                                       " & vbCrLf &
            " INNER JOIN" & vbCrLf &
                "            Produtos ON NFI.Produto_Id = Produtos.Produto_Id ON SubOperacoes.Operacao_Id = NFI.Operacao And" & vbCrLf &
                "            SubOperacoes.SubOperacoes_Id = NFI.SubOperacao RIGHT OUTER JOIN" & vbCrLf &
                "            Comissoes LEFT OUTER JOIN" & vbCrLf &
                "            ContasAPagar Left JOIN" & vbCrLf &
                "            Cotacoes ON ContasAPagar.Baixa = Cotacoes.Data_Id And ContasAPagar.Indexador = Cotacoes.Indexador_Id ON" & vbCrLf &
                "            Comissoes.Empresa_Id = ContasAPagar.EmpresaPedido And Comissoes.EndEmpresa_Id = ContasAPagar.EndEmpresaPedido And" & vbCrLf &
                "            Comissoes.Pedido_Id = ContasAPagar.Pedido ON Pedidos.Pedido_Id = Comissoes.Pedido_Id And Pedidos.Empresa_Id = Comissoes.Empresa_Id And" & vbCrLf &
                "            Pedidos.EndEmpresa_Id = Comissoes.EndEmpresa_Id LEFT OUTER JOIN" & vbCrLf &
                "            Clientes AS Representante ON Comissoes.Representante_Id = Representante.Cliente_Id And" & vbCrLf &
                "            Comissoes.EndRepresentante_Id = Representante.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
                "            Clientes AS Empresa ON Comissoes.Empresa_Id = Empresa.Cliente_Id And Comissoes.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf &
                " WHERE     (Pedidos.Situacao = 1) And  ContasAPagar.Grupado <> 'M'" & vbCrLf &
                "           And (ContasAPagar.Baixa  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

            If ddlRepresentante.SelectedIndex > 0 Then
                Sql &= "        And Comissoes.Representante_Id = '" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            Sql &= " GROUP BY Comissoes.Empresa_Id, Comissoes.EndEmpresa_Id, Empresa.Reduzido, Empresa.Nome, Empresa.Cidade, Empresa.Estado, Comissoes.Representante_Id," & vbCrLf &
                "            Comissoes.EndRepresentante_Id, Representante.Nome, Representante.Cidade, Representante.Estado, Comissoes.Pedido_Id, Comissoes.Percentual," & vbCrLf &
                "            Comissoes.ValorComissao, Comissoes.Principal, Pedidos.DataPedido, ContasAPagar.Registro_Id, ContasAPagar.Baixa, Clientes.Cliente_Id, Clientes.Nome," & vbCrLf &
                "            Clientes.Cidade, ContasAPagar.Provisao, ContasAPagar.ValorDoDocumento, Cotacoes.Indexador_Id, Cotacoes.Indice, Pedidos.Operacao," & vbCrLf &
                "            ContasAPagar.Carteira, NFI.Nota_id" & vbCrLf &
                " HAVING (ISNULL(SUM(CASE WHEN ContasAPagar.Provisao IN (1, 2) THEN ValorDoDocumento ELSE 0 END), 0) <> 0)" & vbCrLf &
                " ORDER BY DataTitulo" & vbCrLf
        End If

        Return Sql
    End Function

    Private Sub CargaRepresentante()
        Dim ds As DataSet
        Dim strRepre As String = ""
        Dim codRep As String = ""

        Sql = " SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome" & vbCrLf &
            " FROM  Clientes INNER JOIN" & vbCrLf &
            "           Comissoes ON Clientes.Cliente_Id = Comissoes.Representante_Id AND Clientes.Endereco_Id = Comissoes.EndRepresentante_Id" & vbCrLf &
            " GROUP BY Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome" & vbCrLf &
            " ORDER BY Clientes.Nome" & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "Representante")

        ddlRepresentante.Items.Clear()

        For Each drRepres As DataRow In ds.Tables(0).Rows
            strRepre = Funcoes.AlinharEsquerda(drRepres("Nome").ToString(), 50, ".") & " - " & drRepres("Cliente_Id") & "-" & drRepres("Endereco_Id")
            codRep = drRepres("Cliente_Id").ToString() & "-" & drRepres("Endereco_Id")
            ddlRepresentante.Items.Add(New ListItem(strRepre, codRep))
        Next

        ddlRepresentante.Items.Insert(0, "")
    End Sub

    Private Function getDataSetExcel() As DataSet
        Dim ds As DataSet = Banco.ConsultaDataSet(GetSql(), "Comissoes")

        Return ds
    End Function

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizaPosicaoDeComissoes()
        Dim Empresa() As String
        Empresa = ddlEmpresa.SelectedValue.Split("-")
        Dim datas As String = Format(CDate(txtDataInicial.Text), "yyyy/MM/dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")
        Dim sqls As New ArrayList()

        If FinanceiroNovo Then

        End If

        Sql = "       Delete    PosicaoDeComissoes;" & vbCrLf &
            " INSERT    INTO PosicaoDeComissoes" & vbCrLf &
            "           (Representante_Id, EndRepresentante_Id, Pedido_Id)" & vbCrLf &
            " SELECT Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id" & vbCrLf &
            " FROM   Pedidos RIGHT OUTER JOIN" & vbCrLf &
            "           ContasAReceber RIGHT OUTER JOIN" & vbCrLf &
            "           Comissoes ON ContasAReceber.EmpresaPedido = Comissoes.Empresa_Id AND ContasAReceber.EndEmpresaPedido = Comissoes.EndEmpresa_Id AND" & vbCrLf &
            "           ContasAReceber.Pedido = Comissoes.Pedido_Id ON Pedidos.Pedido_Id = Comissoes.Pedido_Id AND Pedidos.Empresa_Id = Comissoes.Empresa_Id AND" & vbCrLf &
            "           Pedidos.EndEmpresa_Id = Comissoes.EndEmpresa_Id LEFT OUTER JOIN" & vbCrLf &
            "           Clientes AS Representante ON Comissoes.Representante_Id = Representante.Cliente_Id AND" & vbCrLf &
            "           Comissoes.EndRepresentante_Id = Representante.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
            "           Clientes AS Empresa ON Comissoes.Empresa_Id = Empresa.Cliente_Id AND Comissoes.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf &
            " WHERE (Pedidos.Situacao = 1)" & vbCrLf &
            "   And (Pedidos.DataPedido between '" & datas & "') " & vbCrLf &
            " GROUP BY Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id; " & vbCrLf
        sqls.Add(Sql)

        Sql = " INSERT    INTO PosicaoDeComissoes" & vbCrLf &
            "           (Representante_Id, EndRepresentante_Id, Pedido_Id)" & vbCrLf &
            " SELECT Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, 0 as Pedido_Id" & vbCrLf &
          " FROM   Pedidos RIGHT OUTER JOIN" & vbCrLf &
          "           ContasAReceber RIGHT OUTER JOIN" & vbCrLf &
          "           Comissoes ON ContasAReceber.EmpresaPedido = Comissoes.Empresa_Id AND ContasAReceber.EndEmpresaPedido = Comissoes.EndEmpresa_Id AND" & vbCrLf &
          "           ContasAReceber.Pedido = Comissoes.Pedido_Id ON Pedidos.Pedido_Id = Comissoes.Pedido_Id AND Pedidos.Empresa_Id = Comissoes.Empresa_Id AND" & vbCrLf &
          "           Pedidos.EndEmpresa_Id = Comissoes.EndEmpresa_Id LEFT OUTER JOIN" & vbCrLf &
          "           Clientes AS Representante ON Comissoes.Representante_Id = Representante.Cliente_Id AND" & vbCrLf &
          "           Comissoes.EndRepresentante_Id = Representante.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
          "           Clientes AS Empresa ON Comissoes.Empresa_Id = Empresa.Cliente_Id AND Comissoes.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf &
          " WHERE (Pedidos.Situacao = 1)" & vbCrLf &
          "   And (Pedidos.DataPedido between '" & datas & "') " & vbCrLf &
          " GROUP BY Comissoes.Representante_Id, Comissoes.EndRepresentante_Id " & vbCrLf
        sqls.Add(Sql)

        '---COMISSÕES LEVANTADAS SOBRE AS MERCADORIAS ENTREGUES - Revisado Neri - 19-05-2014
        Sql = " UPDATE PosicaoDeComissoes" & vbCrLf &
                " Set Comissao = Consulta2.Comissao" & vbCrLf &
                " FROM  (SELECT CNPJRepresentante, EndRepresentante, Pedido, SUM(Comissao) AS Comissao" & vbCrLf &
                " FROM   (SELECT CNPJRepresentante, EndRepresentante, Pedido, ValorFaturado, Percentual, ValorFaturado * Percentual / 100 AS Comissao" & vbCrLf &
                "               FROM   (SELECT Comissoes.Representante_Id AS CNPJRepresentante, Comissoes.EndRepresentante_Id AS EndRepresentante, " & vbCrLf &
                "                                              Pedidos.Pedido_Id AS Pedido, CASE WHEN SubOperacoes.Devolucao = 'S' THEN SUM(NotasFiscaisXItens.Valor * - 1) " & vbCrLf &
                "                                              ELSE SUM(NotasFiscaisXItens.Valor) END AS ValorFaturado, Comissoes.Percentual" & vbCrLf &
                "                               FROM   Pedidos INNER JOIN" & vbCrLf &
                "                                              NotasFiscaisXItens ON Pedidos.Pedido_Id = NotasFiscaisXItens.Pedido AND Pedidos.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND" & vbCrLf &
                "                                              Pedidos.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id INNER JOIN" & vbCrLf &
                "                                              SubOperacoes ON NotasFiscaisXItens.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf &
                "                                              NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
                "                                              NotasFiscais ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id AND" & vbCrLf &
                " NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id And " & vbCrLf &
                " NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id And " & vbCrLf &
                " NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id And " & vbCrLf &
                "                                               NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id RIGHT OUTER JOIN" & vbCrLf &
                "                                               Comissoes ON Pedidos.Pedido_Id = Comissoes.Pedido_Id AND Pedidos.Empresa_Id = Comissoes.Empresa_Id AND " & vbCrLf &
                "                                               Pedidos.EndEmpresa_Id = Comissoes.EndEmpresa_Id" & vbCrLf &
                " WHERE (Pedidos.Situacao = 1) And Pedidos.Operacao <> 36 " & vbCrLf &
                " And (NotasFiscais.Situacao = 1) And SubOperacoes.Classe not In ('" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf &
                " And (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
                "                               GROUP BY Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id, Comissoes.Percentual, " & vbCrLf &
                "                                              Comissoes.ValorComissao, Comissoes.Principal, SubOperacoes.Devolucao) AS Consulta) AS Consulta1" & vbCrLf &
                " GROUP BY CNPJRepresentante, EndRepresentante, Pedido) AS Consulta2 INNER JOIN" & vbCrLf &
                " PosicaoDeComissoes ON Consulta2.CNPJRepresentante = PosicaoDeComissoes.Representante_Id AND " & vbCrLf &
                " Consulta2.EndRepresentante = PosicaoDeComissoes.EndRepresentante_Id And Consulta2.Pedido = PosicaoDeComissoes.Pedido_Id"
        sqls.Add(Sql)

        '---COMISSÕES LIBERADAS PELO PAGAMENTO DOS CLIENTES  - Revisado Neri - 19-05-2014
        Sql = "  Update PosicaoDeComissoes" & vbCrLf &
            " Set Liberado = Consulta.Valor" & vbCrLf &
            " FROM  (SELECT Representante_Id, EndRepresentante_Id, Pedido, SUM(Liberado) AS Valor" & vbCrLf &
            "           FROM   (SELECT Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id As Pedido, ContasAReceber.Registro_Id, " & vbCrLf &
            "                          ContasAReceber.ValorDoDocumento * Comissoes.Percentual / 100 AS Liberado" & vbCrLf &
            "                           FROM   ContasAReceber RIGHT OUTER JOIN" & vbCrLf &
            "                                          PosicaoDeComissoes INNER JOIN" & vbCrLf &
            "                                          Comissoes ON PosicaoDeComissoes.Representante_Id = Comissoes.Representante_Id AND" & vbCrLf &
            "                                          PosicaoDeComissoes.EndRepresentante_Id = Comissoes.EndRepresentante_Id ON" & vbCrLf &
            "                                          ContasAReceber.EmpresaPedido = Comissoes.Empresa_Id AND ContasAReceber.EndEmpresaPedido = Comissoes.EndEmpresa_Id AND" & vbCrLf &
            "                                          ContasAReceber.Pedido = Comissoes.Pedido_Id LEFT OUTER JOIN" & vbCrLf &
            "                                          Pedidos INNER JOIN" & vbCrLf &
            "                                          NotasFiscaisXItens ON Pedidos.Pedido_Id = NotasFiscaisXItens.Pedido And Pedidos.Empresa_Id = NotasFiscaisXItens.Empresa_Id And" & vbCrLf &
            "                                          Pedidos.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id ON Comissoes.Pedido_Id = Pedidos.Pedido_Id AND" & vbCrLf &
            "                                          Comissoes.Empresa_Id = Pedidos.Empresa_Id And Comissoes.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf &
            "                           WHERE (Pedidos.Situacao = 1) And Pedidos.Operacao <> 36 And (ContasAReceber.Provisao = 1) And Grupado <> 'M' " & vbCrLf &
            "                             And (ContasAReceber.Baixa between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "                           GROUP BY Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id, ContasAReceber.Registro_Id, ContasAReceber.ValorDoDocumento, Comissoes.Percentual, ContasAReceber.Carteira)" & vbCrLf &
            "                          AS Liberados" & vbCrLf &
            "           GROUP BY Representante_Id, EndRepresentante_Id, Pedido) AS Consulta INNER JOIN" & vbCrLf &
            "           PosicaoDeComissoes ON Consulta.Representante_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "           Consulta.EndRepresentante_Id = PosicaoDeComissoes.EndRepresentante_Id" & vbCrLf &
            "       And Consulta.Pedido = PosicaoDeComissoes.Pedido_Id"
        sqls.Add(Sql)

        '---COMISSÕES LIBERADAS PELO RECEBIMENTO DOS CLIENTES  - Revisado Neri - 19-05-2014
        Sql = "  Update PosicaoDeComissoes" & vbCrLf &
        "       Set Liberado = Liberado - Consulta.Valor" & vbCrLf &
        "       FROM  (SELECT Representante_Id, EndRepresentante_Id, Pedido, SUM(Liberado) AS Valor" & vbCrLf &
        "           FROM   (SELECT Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id As Pedido, ContasAPagar.Registro_Id, " & vbCrLf &
        "                          ContasAPagar.ValorDoDocumento * Comissoes.Percentual / 100 AS Liberado" & vbCrLf &
        "                           FROM   ContasAPagar RIGHT OUTER JOIN" & vbCrLf &
        "                                          PosicaoDeComissoes INNER JOIN" & vbCrLf &
        "                                          Comissoes ON PosicaoDeComissoes.Representante_Id = Comissoes.Representante_Id AND" & vbCrLf &
        "                                          PosicaoDeComissoes.EndRepresentante_Id = Comissoes.EndRepresentante_Id ON" & vbCrLf &
        "                                          ContasAPagar.EmpresaPedido = Comissoes.Empresa_Id AND ContasAPagar.EndEmpresaPedido = Comissoes.EndEmpresa_Id AND" & vbCrLf &
        "                                          ContasAPagar.Pedido = Comissoes.Pedido_Id LEFT OUTER JOIN" & vbCrLf &
        "                                          Pedidos INNER JOIN" & vbCrLf &
        "                                          NotasFiscaisXItens ON Pedidos.Pedido_Id = NotasFiscaisXItens.Pedido And Pedidos.Empresa_Id = NotasFiscaisXItens.Empresa_Id And " & vbCrLf &
        "                                          Pedidos.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id ON Comissoes.Pedido_Id = Pedidos.Pedido_Id AND" & vbCrLf &
        "                                          Comissoes.Empresa_Id = Pedidos.Empresa_Id And Comissoes.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf &
        "                            WHERE (Pedidos.Situacao = 1) And Pedidos.Operacao <> 36 And (ContasAPagar.Provisao = 1) And Grupado <> 'M'" & vbCrLf &
        "                                     And (ContasAPagar.Baixa between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
        "                           GROUP BY Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id, ContasAPagar.Registro_Id, ContasAPagar.ValorDoDocumento, Comissoes.Percentual, ContasAPagar.Carteira)" & vbCrLf &
        "                          AS Liberados" & vbCrLf &
        "           GROUP BY Representante_Id, EndRepresentante_Id, Pedido) AS Consulta INNER JOIN" & vbCrLf &
        "           PosicaoDeComissoes ON Consulta.Representante_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
        "           Consulta.EndRepresentante_Id = PosicaoDeComissoes.EndRepresentante_Id" & vbCrLf &
        "       And Consulta.Pedido = PosicaoDeComissoes.Pedido_Id"
        sqls.Add(Sql)

        '-------------------------LIBERADO PELA ENTREGA DO PRODUTO DE COMPRAS
        Sql = "   Update PosicaoDeComissoes" & vbCrLf &
            " Set Liberado = PosicaoDeComissoes.Liberado + Consulta2.Liberado" & vbCrLf &
            " FROM  (SELECT CNPJRepresentante, EndRepresentante, Pedido, SUM(Liberado) AS Liberado" & vbCrLf &
            "          FROM   (SELECT CNPJRepresentante, EndRepresentante, Pedido, ValorFaturado, Percentual, ValorFaturado * Percentual / 100 AS Liberado" & vbCrLf &
            "                          FROM   (SELECT Comissoes.Representante_Id AS CNPJRepresentante, Comissoes.EndRepresentante_Id AS EndRepresentante," & vbCrLf &
            "                                                         Pedidos.Pedido_Id AS Pedido, CASE WHEN SubOperacoes.Devolucao = 'S' THEN SUM(NotasFiscaisXItens.Valor * - 1)" & vbCrLf &
            "                                                         ELSE SUM(NotasFiscaisXItens.Valor) END AS ValorFaturado, Comissoes.Percentual" & vbCrLf &
            "                                          FROM   Pedidos INNER JOIN" & vbCrLf &
            "                                                        NotasFiscaisXItens ON Pedidos.Pedido_Id = NotasFiscaisXItens.Pedido INNER JOIN" & vbCrLf &
            "                                                        SubOperacoes ON NotasFiscaisXItens.Operacao = SubOperacoes.Operacao_Id AND" & vbCrLf &
            "                                                         NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
            "                                                         NotasFiscais ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id AND" & vbCrLf &
            "                                                         NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id AND" & vbCrLf &
            "                                                         NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id AND" & vbCrLf &
            "                                                         NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id AND" & vbCrLf &
            "                                                         NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id RIGHT OUTER JOIN" & vbCrLf &
            "                                                         Comissoes ON Pedidos.Pedido_Id = Comissoes.Pedido_Id AND Pedidos.Empresa_Id = Comissoes.Empresa_Id AND" & vbCrLf &
            "                                                         Pedidos.EndEmpresa_Id = Comissoes.EndEmpresa_Id" & vbCrLf &
            "                                          WHERE (Pedidos.Situacao = 1) And Pedidos.Operacao <> 36 AND (NotasFiscais.Situacao = 1) " & vbCrLf &
            "                                           And (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "                                           And (SubOperacoes.Classe IN ('" & eClassesOperacoes.COMPRAS.ToString & "', '" & eClassesOperacoes.COMPRASAORDEM.ToString & "'))" & vbCrLf &
            "                                           GROUP BY Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Pedidos.Pedido_Id, Comissoes.Percentual," & vbCrLf &
            "                                                          Comissoes.ValorComissao, Comissoes.Principal, SubOperacoes.Devolucao) AS Consulta) AS Consulta1" & vbCrLf &
            "          GROUP BY CNPJRepresentante, EndRepresentante, Pedido) AS Consulta2 INNER JOIN" & vbCrLf &
            "           PosicaoDeComissoes ON Consulta2.CNPJRepresentante = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "           Consulta2.EndRepresentante = PosicaoDeComissoes.EndRepresentante_Id And Consulta2.Pedido = PosicaoDeComissoes.Pedido_Id" & vbCrLf
        sqls.Add(Sql)

        '----- Ajusta Liberaçao a maior para limitar ao Entregue
        Sql = " Update posicaodecomissoes" & vbCrLf &
                " Set Liberado = Comissao" & vbCrLf &
                " where Liberado > Comissao"
        sqls.Add(Sql)

        '--- FATURADO PELOS REPRESENTANTES
        'Neri
        Sql = " Update PosicaoDeComissoes" & vbCrLf &
            " Set Faturado = Consulta.Faturado" & vbCrLf &
            " FROM  (SELECT PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id, SUM(NotasFiscaisXItens.Valor) AS Faturado" & vbCrLf &
            "            FROM   NotasFiscaisXItens INNER JOIN" & vbCrLf &
            "                           NotasFiscais ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id AND NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id AND NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id INNER JOIN" & vbCrLf &
            "                           PosicaoDeComissoes ON NotasFiscais.Cliente_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "                           NotasFiscais.EndCliente_Id = PosicaoDeComissoes.EndRepresentante_Id  And PosicaoDeComissoes.Pedido_Id = 0" & vbCrLf &
            "            WHERE (NotasFiscais.Situacao = 1)" & vbCrLf &
            "              And (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "            GROUP BY PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id) AS Consulta INNER JOIN" & vbCrLf &
            "            PosicaoDeComissoes ON Consulta.Representante_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "            Consulta.EndRepresentante_Id = PosicaoDeComissoes.EndRepresentante_Id" & vbCrLf &
            "            And 0 = PosicaoDeComissoes.Pedido_Id" & vbCrLf
        sqls.Add(Sql)

        '--- PAGO PARA O REPRESENTANTES
        Sql = " Update PosicaoDeComissoes" & vbCrLf &
            " Set Pago = Consulta.Pago" & vbCrLf &
            " FROM  (SELECT PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id, SUM(ContasAPagar.ValorDoDocumento) AS Pago" & vbCrLf &
            "           FROM   PosicaoDeComissoes LEFT OUTER JOIN" & vbCrLf &
            "                          ContasAPagar ON PosicaoDeComissoes.Representante_Id = ContasAPagar.Cliente AND" & vbCrLf &
            "                          PosicaoDeComissoes.EndRepresentante_Id = ContasAPagar.EndCliente  And PosicaoDeComissoes.Pedido_Id = 0" & vbCrLf &
            "           WHERE (ContasAPagar.Situacao = 1) And (ContasAPagar.Provisao = 1) And Grupado <> 'M' And (ContasAPagar.ContaContabilCliente <> '2010112')" & vbCrLf &
            "             And (ContasAPagar.Baixa between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "           GROUP BY PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id) AS Consulta INNER JOIN" & vbCrLf &
            "           PosicaoDeComissoes ON Consulta.Representante_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "           Consulta.EndRepresentante_Id = PosicaoDeComissoes.EndRepresentante_Id" & vbCrLf &
            "           And 0 = PosicaoDeComissoes.Pedido_Id" & vbCrLf
        sqls.Add(Sql)

        '--- DEVOLUCAO DE PAGAMENTO DO REPRESENTANTES
        Sql = " Update PosicaoDeComissoes" & vbCrLf &
            " Set PosicaoDeComissoes.Pago = PosicaoDeComissoes.Pago - Consulta.Pago" & vbCrLf &
            " FROM  (SELECT PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id, SUM(ContasAReceber.ValorDoDocumento) AS Pago" & vbCrLf &
            "           FROM   PosicaoDeComissoes LEFT OUTER JOIN" & vbCrLf &
            "                          ContasAReceber ON PosicaoDeComissoes.Representante_Id = ContasAReceber.Cliente AND" & vbCrLf &
            "                          PosicaoDeComissoes.EndRepresentante_Id = ContasAReceber.EndCliente  And PosicaoDeComissoes.Pedido_Id = 0" & vbCrLf &
            "           WHERE (ContasAReceber.Situacao = 1) And (ContasAReceber.Provisao = 1) And Grupado <> 'M'" & vbCrLf &
            "             And (ContasAReceber.Baixa between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "           GROUP BY PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id) AS Consulta INNER JOIN" & vbCrLf &
            "           PosicaoDeComissoes ON Consulta.Representante_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "           Consulta.EndRepresentante_Id = PosicaoDeComissoes.EndRepresentante_Id" & vbCrLf &
            "           And 0 = PosicaoDeComissoes.Pedido_Id" & vbCrLf
        sqls.Add(Sql)

        '--- FATURADO PELOS REPRESENTANTES
        Sql = "  Update PosicaoDeComissoes" & vbCrLf &
            "  Set Encargo = Consulta.IRRF" & vbCrLf &
            " FROM  (SELECT PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id, SUM(NotasFiscaisXEncargos.Valor) AS IRRF" & vbCrLf &
            "            FROM   NotasFiscaisXItens INNER JOIN" & vbCrLf &
            "                           NotasFiscais ON NotasFiscaisXItens.Empresa_Id = NotasFiscais.Empresa_Id AND NotasFiscaisXItens.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.Cliente_Id = NotasFiscais.Cliente_Id AND NotasFiscaisXItens.EndCliente_Id = NotasFiscais.EndCliente_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscais.Serie_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.Nota_Id = NotasFiscais.Nota_Id INNER JOIN" & vbCrLf &
            "                           PosicaoDeComissoes ON NotasFiscais.Cliente_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "                           NotasFiscais.EndCliente_Id = PosicaoDeComissoes.EndRepresentante_Id  And PosicaoDeComissoes.Pedido_Id = 0 INNER JOIN" & vbCrLf &
            "                           NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND" & vbCrLf &
            "                          NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf &
            "                           NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id AND NotasFiscaisXEncargos.Encargo_Id = 'IRRF PJ'" & vbCrLf &
            "            WHERE (NotasFiscais.Situacao = 1)" & vbCrLf &
            "              And (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "            GROUP BY PosicaoDeComissoes.Representante_Id, PosicaoDeComissoes.EndRepresentante_Id) AS Consulta INNER JOIN" & vbCrLf &
            "            PosicaoDeComissoes ON Consulta.Representante_Id = PosicaoDeComissoes.Representante_Id AND" & vbCrLf &
            "            Consulta.EndRepresentante_Id = PosicaoDeComissoes.EndRepresentante_Id" & vbCrLf &
            "            And 0 = PosicaoDeComissoes.Pedido_Id" & vbCrLf
        sqls.Add(Sql)

        'COMISSÃO LIBELADA PELO RECEBIMENTO DO CÂMBIO
        Sql = "       Update PosicaoDeComissoes" & vbCrLf &
            " SET   Comissao =  Final.Liberado, " & vbCrLf &
            "       Liberado =  Final.Liberado" & vbCrLf &
            " FROM  (SELECT Pedido_Id AS Pedido, SUM((Quantidade * Percentual * Indice) * (ValorDoDocumento / Valor * 100) / 100) AS Liberado" & vbCrLf &
            "            FROM   (SELECT Consulta.Pedido_Id, Consulta.Quantidade, Consulta.Valor, ContasAReceber.Registro_Id, ContasAReceber.ValorDoDocumento, Cotacoes.Indice," & vbCrLf &
            "                                           Comissoes.Percentual" & vbCrLf &
            "                            FROM   Comissoes RIGHT OUTER JOIN" & vbCrLf &
            "                                               (SELECT Pedidos.Pedido_Id, SUM(NotasFiscaisXItens.QuantidadeFiscal / 1000) AS Quantidade, SUM(NotasFiscaisXItens.Valor) AS Valor" & vbCrLf &
            "                                                FROM   Pedidos LEFT OUTER JOIN" & vbCrLf &
            "                                                               NotasFiscaisXItens ON Pedidos.Pedido_Id = NotasFiscaisXItens.Pedido" & vbCrLf &
            "                                                WHERE (Pedidos.Situacao = 1) And (Pedidos.Operacao = 36) And (NotasFiscaisXItens.QuantidadeFiscal > 0)" & vbCrLf &
            "                                                GROUP BY Pedidos.Pedido_Id) AS Consulta ON Comissoes.Pedido_Id = Consulta.Pedido_Id LEFT OUTER JOIN" & vbCrLf &
            "                                           Cotacoes INNER JOIN" & vbCrLf &
            "                                           ContasAReceber ON Cotacoes.Indexador_Id = ContasAReceber.Indexador AND Cotacoes.Data_Id = ContasAReceber.Baixa ON" & vbCrLf &
            "                                           Consulta.Pedido_Id = ContasAReceber.Pedido" & vbCrLf &
            "                            WHERE (ContasAReceber.Registro_Id > 0) AND (ContasAReceber.Provisao = 1) AND (ContasAReceber.Grupado <> 'M') AND (Comissoes.Percentual > 0)" & vbCrLf &
            "                              And (ContasAReceber.Baixa between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "            ) AS Total" & vbCrLf &
            "            GROUP BY Pedido_Id) AS Final INNER JOIN" & vbCrLf &
            "            PosicaoDeComissoes ON Final.Pedido = PosicaoDeComissoes.Pedido_Id" & vbCrLf
        sqls.Add(Sql)

        Try
            Banco.GravaBanco(sqls)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeComissoesPedido", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa, True)
                    CargaRepresentante()
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gerencial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub AtualizaPosicaoTable()

        '************PARTE 1 - LINHA 775****************************
        Sql = "SELECT p.Empresa_Id," & vbCrLf &
                "       p.EndEmpresa_Id," & vbCrLf &
                "       p.Pedido_Id," & vbCrLf &
                "       p.OrigemDestino," & vbCrLf &
                "       p.Moeda AS MoedaPedido," & vbCrLf &
                "       sOp.EntradaSaida," & vbCrLf &
                "       pXi.Produto_Id           AS Produto," & vbCrLf &
                "       convert(Decimal(18,4),0) AS Contratado," & vbCrLf &
                "       convert(Decimal(18,4),0) AS Laudo," & vbCrLf &
                "       convert(Decimal(18,4),0) AS Entregue," & vbCrLf &
                "	    convert(Decimal(18,4),0) AS AEntregar," & vbCrLf &
                "	    convert(Decimal(18,4),0) AS QuantidadeFixado," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS ValorFixadoOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS ValorFixadoMoeda," & vbCrLf &
                "	    convert(Decimal(18,4),0) AS AFixar," & vbCrLf &
                "	    convert(Decimal(18,4),0) AS Pago," & vbCrLf &
                "	    convert(Decimal(18,4),0) AS PagoNaoRecebido," & vbCrLf &
                "	    convert(Decimal(18,4),0) AS RecebidoNaoPago," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Adiantamento," & vbCrLf &
                "	    convert(Decimal(18,10),0)AS UnitarioOficial," & vbCrLf &
                "	    convert(Decimal(18,10),0)AS UnitarioMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS TotalOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS TotalMoeda," & vbCrLf &
                "       convert(Decimal(18,2),0) AS TotalLiquidoOficial," & vbCrLf &
                "       convert(Decimal(18,2),0) AS TotalLiquidoMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Programado_PagarOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Baixado_PagarOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Programado_PagarMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Baixado_PagarMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Programado_ReceberOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Baixado_ReceberOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Programado_ReceberMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Baixado_ReceberMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Programado_SaldoOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Baixado_SaldoOficial," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Programado_SaldoMoeda," & vbCrLf &
                "	    convert(Decimal(18,2),0) AS Baixado_SaldoMoeda," & vbCrLf &
                "       convert(Decimal(18,2),0) AS ValorLiquidoDeNota," & vbCrLf &
                "       convert(Decimal(18,2),0) AS Complementacoes" & vbCrLf &
                "  into #Posicao" & vbCrLf &
                "  FROM Pedidos p" & vbCrLf &
                " INNER JOIN PedidoXItem pXi" & vbCrLf &
                "    ON pXi.Empresa_Id    = p.Empresa_Id" & vbCrLf &
                "   AND pXi.EndEmpresa_Id = p.EndEmpresa_Id" & vbCrLf &
                "   AND pXi.Pedido_Id     = p.Pedido_Id" & vbCrLf &
                " INNER JOIN Produtos Prd" & vbCrLf &
                "    ON Prd.Produto_Id = pXi.Produto_Id" & vbCrLf &
                " INNER JOIN SubOperacoes sOp" & vbCrLf &
                " 	 ON sOp.Operacao_Id     = p.Operacao" & vbCrLf &
                "   AND sOp.SubOperacoes_Id = p.SubOperacao" & vbCrLf &
                " WHERE p.Empresa_Id    = " & ddlEmpresa.SelectedValue.Split("-")(0) & vbCrLf &
                "   AND p.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                "   And Prd.Agrupar     ='N'" & vbCrLf &
                "   And p.Situacao      = 1" & vbCrLf &
                "   And p.DataPedido between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf

        'Atualiza Quantidade Contratada Itens do Pedido----
        'PARTE 2 - LINHA 798
        Sql &= "Update #Posicao SET                                                                  " & vbCrLf &
                "   #Posicao.Contratado      = consulta.Quantidade,                                   " & vbCrLf &
                "   #Posicao.TotalOficial    = consulta.TotalOficial,                                 " & vbCrLf &
                "   #Posicao.TotalMoeda      = consulta.TotalMoeda,                                    " & vbCrLf &
                "   #Posicao.TotalLiquidoOficial = consulta.LiqOficial,                                 " & vbCrLf &
                "   #Posicao.TotalLiquidoMoeda   = consulta.LiqMoeda  " & vbCrLf &
                "   FROM (SELECT pXi.Empresa_Id,                                                      " & vbCrLf &
                "                pXi.EndEmpresa_Id,                                                   " & vbCrLf &
                "                pXi.Pedido_Id,                                                       " & vbCrLf &
                "                pXi.Produto_Id,                                                      " & vbCrLf &
                "                                                                                     " & vbCrLf &
                "               encLiq.ValorOficial as LiqOficial," & vbCrLf &
                "               encLiq.ValorMoeda   as LiqMoeda,  " & vbCrLf &
                "                                                                                       " & vbCrLf &
                "			    SUM(case                                                              " & vbCrLf &
                "			          when sOp.Classe <> '" & eClassesOperacoes.AFIXAR.ToString & "'                                     " & vbCrLf &
                "			            THEN CASE                                                     " & vbCrLf &
                "			                   WHEN pXi.TipoDeLancamento = 'E'                        " & vbCrLf &
                "			                     THEN pXi.TotalOficial * - 1                          " & vbCrLf &
                "			                     ELSE pXi.TotalOficial                                " & vbCrLf &
                "        End                                                                          " & vbCrLf &
                "			            Else 0                                                        " & vbCrLf &
                "			        End)  AS TotalOficial,                                            " & vbCrLf &
                "                                                                                     " & vbCrLf &
                "			    SUM(case                                                              " & vbCrLf &
                "			          when sOp.Classe <> '" & eClassesOperacoes.AFIXAR.ToString & "'                                     " & vbCrLf &
                "			            THEN CASE                                                     " & vbCrLf &
                "			                   WHEN pXi.TipoDeLancamento = 'E'                        " & vbCrLf &
                "			                     THEN pXi.TotalMoeda * - 1                            " & vbCrLf &
                "			                     ELSE pXi.TotalMoeda                                  " & vbCrLf &
                "        End                                                                          " & vbCrLf &
                "			            Else 0                                                        " & vbCrLf &
                "			        End) AS TotalMoeda,                                               " & vbCrLf &
                "                                                                                     " & vbCrLf &
                "			    SUM(CASE                                                              " & vbCrLf &
                "			          WHEN pXi.TipoDeLancamento = 'E'                                 " & vbCrLf &
                "			            THEN pXi.Quantidade * - 1                                     " & vbCrLf &
                "			            ELSE pXi.Quantidade                                           " & vbCrLf &
                "			        END) AS Quantidade                                                " & vbCrLf &
                "                                                                                     " & vbCrLf &
                "			FROM PedidoXItemxLancamento pXi                                                    " & vbCrLf &
                "		   INNER JOIN Pedidos p                                                       " & vbCrLf &
                "		      ON p.Empresa_Id    = pXi.Empresa_Id                                     " & vbCrLf &
                "			 AND p.EndEmpresa_Id = pXi.EndEmpresa_Id                                  " & vbCrLf &
                "			 AND p.Pedido_Id     = pXi.Pedido_Id                                      " & vbCrLf &
                "		   INNER JOIN SubOperacoes sOp                                                " & vbCrLf &
                "		      ON sOp.Operacao_Id     = p.Operacao                                     " & vbCrLf &
                "			 AND sOp.SubOperacoes_Id = p.SubOperacao                                  " & vbCrLf &
                "                                                                                       " & vbCrLf &
                "           Inner Join(Select   Empresa_id," & vbCrLf &
                "                               EndEmpresa_id,                                   " & vbCrLf &
                "                               Pedido_Id,                                        " & vbCrLf &
                "                               Produto_id,                                       " & vbCrLf &
                "                               ValorOficial,                                     " & vbCrLf &
                "                               ValorMoeda                                        " & vbCrLf &
                "               from Pedidosxencargos                            " & vbCrLf &
                "               where Encargo_id = 'LIQUIDO' 	" & vbCrLf &
                "              ) As EncLiq                         " & vbCrLf &
                "           ON pXi.Empresa_Id    = EncLiq.Empresa_Id    " & vbCrLf &
                "	        AND pXi.EndEmpresa_Id = EncLiq.EndEmpresa_Id  " & vbCrLf &
                "	        AND pXi.Pedido_Id     = EncLiq.Pedido_Id      " & vbCrLf &
                "	        AND pXi.Produto_id    = EncLiq.Produto_id     " & vbCrLf &
                "                                                           " & vbCrLf &
                "		   WHERE pXi.Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf &
                "		     AND pXi.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & "   " & vbCrLf &
                "		     And p.Situacao        = 1                                                " & vbCrLf &
                "		   GROUP BY pXi.Empresa_Id, pXi.EndEmpresa_Id, pXi.Pedido_Id, pXi.Produto_Id, encLiq.ValorOficial, encLiq.ValorMoeda  " & vbCrLf &
                "		 ) AS consulta                                                                " & vbCrLf &
                "  INNER JOIN #Posicao                                                                " & vbCrLf &
                "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id                                  " & vbCrLf &
                "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id                               " & vbCrLf &
                "	AND	consulta.Pedido_Id     = #Posicao.Pedido_Id                                   " & vbCrLf &
                "	AND	consulta.Produto_Id    = #Posicao.Produto                                     " & vbCrLf

        'Atualiza Complementações de Preco ---------
        'PARTE 3 - LINHA 833
        Sql &= "Update #Posicao SET                                                                       " & vbCrLf &
           "	#Posicao.Complementacoes = consulta.Quantidade,                                        " & vbCrLf &
           "	#Posicao.TotalOficial    = consulta.TotalOficial,                                      " & vbCrLf &
           "	#Posicao.TotalMoeda      = consulta.TotalMoeda                                         " & vbCrLf &
           "  FROM (SELECT p.Empresa_Id,                                                            " & vbCrLf &
           "               p.EndEmpresa_Id,                                                         " & vbCrLf &
           "               p.Pedido_Id,                                                             " & vbCrLf &
           "               pXiXf.Produto_Id,                                                            " & vbCrLf &
           "			   SUM(CASE                                                                    " & vbCrLf &
           "			         WHEN sOp.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'                                   " & vbCrLf &
           "			           THEN pXiXf.TotalOficial                                             " & vbCrLf &
           "			           ELSE 0.00                                                           " & vbCrLf &
           "			       END) AS TotalOficial,                                                   " & vbCrLf &
           "			   SUM(CASE                                                                    " & vbCrLf &
           "			         WHEN sOp.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'                                   " & vbCrLf &
           "			           THEN pXiXf.TotalMoeda                                               " & vbCrLf &
           "			           ELSE 0.00                                                           " & vbCrLf &
           "			       END) AS TotalMoeda,                                                     " & vbCrLf &
           "			   SUM(CASE                                                                    " & vbCrLf &
           "			         WHEN sOp.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'                                   " & vbCrLf &
           "			           THEN pXiXf.Quantidade                                               " & vbCrLf &
           "			           ELSE 0.00                                                           " & vbCrLf &
           "			       END) AS Quantidade                                                      " & vbCrLf &
           "		  FROM Pedidos p                                                              " & vbCrLf &
           "	     INNER JOIN VW_PedidosXItensXFixacoes pXiXf                                           " & vbCrLf &
           "			ON pXiXf.Empresa_Id    = p.Empresa_Id                                        " & vbCrLf &
           "		   AND pXiXf.EndEmpresa_Id = p.EndEmpresa_Id                                     " & vbCrLf &
           "		   AND pXiXf.Pedido_Id     = p.Pedido_Id                                         " & vbCrLf &
           "	     INNER JOIN SubOperacoes sOp                                                       " & vbCrLf &
           "			ON sOp.Operacao_Id     = pXiXf.Operacao                                        " & vbCrLf &
           "		   AND sOp.SubOperacoes_Id = pXiXf.SubOperacao                                     " & vbCrLf &
           "		 WHERE p.Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'        " & vbCrLf &
           "		   AND p.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & "          " & vbCrLf &
           "		   And p.Situacao        = 1                                                       " & vbCrLf &
           "		 GROUP BY p.Empresa_Id, p.EndEmpresa_Id, p.Pedido_Id, pXiXf.Produto_Id         " & vbCrLf &
           "		) AS consulta                                                                      " & vbCrLf &
           "	INNER JOIN #Posicao                                                                    " & vbCrLf &
           "	   ON consulta.Empresa_Id    = #Posicao.Empresa_Id                                     " & vbCrLf &
           "	  AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id                                  " & vbCrLf &
           "	  AND consulta.Pedido_Id     = #Posicao.Pedido_Id                                      " & vbCrLf &
           "	  AND consulta.Produto_Id    = #Posicao.Produto                                        " & vbCrLf &
           "	INNER JOIN Pedidos AS ConsultasXPedidos                                                " & vbCrLf &
           "	   ON ConsultasXPedidos.Empresa_Id    = #Posicao.Empresa_Id                            " & vbCrLf &
           "	  AND ConsultasXPedidos.EndEmpresa_Id = #Posicao.EndEmpresa_Id                         " & vbCrLf &
           "	  AND ConsultasXPedidos.Pedido_Id     = #Posicao.Pedido_Id                             " & vbCrLf &
           "	INNER JOIN SubOperacoes AS ConsultasXSubOperacoes                                      " & vbCrLf &
           "	   ON ConsultasXSubOperacoes.Operacao_Id      = ConsultasXPedidos.Operacao             " & vbCrLf &
           "	  AND ConsultasXSubOperacoes.SubOperacoes_Id = ConsultasXPedidos.SubOperacao           " & vbCrLf &
           "	WHERE (ConsultasXSubOperacoes.Classe = '" & eClassesOperacoes.AFIXAR.ToString & "')                                       " & vbCrLf

        'Atualiza Quantidade Fixada Pedidos X Itens X Fixacoes ----
        'PARTE 4 - LINHA 867
        Sql &= "Update #Posicao SET                                                        " & vbCrLf &
               "	#Posicao.QuantidadeFixado   = consulta.Quantidade,                      " & vbCrLf &
               "	#Posicao.ValorFixadoOficial = consulta.TotalOficial,                    " & vbCrLf &
               "	#Posicao.ValorFixadoMoeda   = consulta.TotalMoeda                       " & vbCrLf &
               "  FROM (SELECT pXiXf.Empresa_Id,                                           " & vbCrLf &
               "	           pXiXf.EndEmpresa_Id,                                         " & vbCrLf &
               "	           pXiXf.Pedido_Id,                                             " & vbCrLf &
               "			   SUM(pXiXf.Quantidade) AS Quantidade,                         " & vbCrLf &
               "			   SUM(pXiXfXe.ValorOficial) as TotalOficial,                   " & vbCrLf &
               "			   SUM(pXiXfXe.ValorMoeda) As TotalMoeda                        " & vbCrLf &
               "		  FROM VW_PedidosXItensXFixacoes pXiXf                                 " & vbCrLf &
               "		 INNER JOIN VW_PedidosXItensXFixacoesXEncargos pXiXfXe              " & vbCrLf &
               "			ON pXiXfXe.Empresa_Id    = pXiXf.Empresa_Id                     " & vbCrLf &
               "		   AND pXiXfXe.EndEmpresa_Id = pXiXf.EndEmpresa_Id                  " & vbCrLf &
               "		   AND pXiXfXe.Pedido_Id     = pXiXf.Pedido_Id                      " & vbCrLf &
               "		   AND pXiXfXe.Produto_Id    = pXiXf.Produto_Id                     " & vbCrLf &
               "		   AND pXiXfXe.Fixacao_Id    = pXiXf.Fixacao_Id                     " & vbCrLf &
               "		 INNER JOIN #Posicao                                                " & vbCrLf &
               "		    ON #Posicao.Empresa_Id    = pXiXf.Empresa_Id                    " & vbCrLf &
               "		   AND #Posicao.EndEmpresa_Id = pXiXf.EndEmpresa_Id                 " & vbCrLf &
               "		   AND #Posicao.Pedido_Id     = pXiXf.Pedido_Id                     " & vbCrLf &
               "		 WHERE pXiXf.Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'          " & vbCrLf &
               "		   AND pXiXf.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & "                     " & vbCrLf &
               "		   AND pXiXfXe.Encargo_Id  = 'LIQUIDO'                              " & vbCrLf &
               "		 GROUP BY pXiXf.Empresa_Id, pXiXf.EndEmpresa_Id, pXiXf.Pedido_Id    " & vbCrLf &
               "		) AS consulta                                                       " & vbCrLf &
               "  INNER JOIN #Posicao                                                      " & vbCrLf &
               "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id                        " & vbCrLf &
               "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id                     " & vbCrLf &
               "	AND	consulta.Pedido_Id     = #Posicao.Pedido_Id                         " & vbCrLf

        'Atualiza Quantidade Entregue Notas Fiscais X Itens ---
        'PARTE 5 - LINHA 901
        Sql &= "Update #Posicao SET                                                                            " & vbCrLf &
           "	#Posicao.Entregue           = consulta.Entregue,                                            " & vbCrLf &
           "	#Posicao.ValorLiquidoDeNota = consulta.Valor                                                " & vbCrLf &
           "  FROM (SELECT n.Empresa_Id,                                                                   " & vbCrLf &
           "               n.EndEmpresa_Id,                                                                " & vbCrLf &
           "               n.Pedido,                                                                       " & vbCrLf &
           "			   SUM(CASE                                                                         " & vbCrLf &
           "			         WHEN sOP.Devolucao = 'S'                                                   " & vbCrLf &
           "			           THEN nXi.QuantidadeFiscal * - 1                                          " & vbCrLf &
           "			           ELSE nXi.QuantidadeFiscal                                                " & vbCrLf &
           "			       END) AS Entregue,                                                            " & vbCrLf &
           "			   SUM(CASE                                                                         " & vbCrLf &
           "			         WHEN sOP.Devolucao = 'S'                                                   " & vbCrLf &
           "			           THEN nxE.Valor * - 1                                                     " & vbCrLf &
           "			           ELSE nXe.Valor                                                           " & vbCrLf &
           "			       END) AS Valor                                                                " & vbCrLf &
           "		  FROM NotasFiscaisXItens nXi                                                           " & vbCrLf &
           "		 INNER JOIN NotasFiscais n                                                              " & vbCrLf &
           "			ON n.Empresa_Id      = nXi.Empresa_Id                                               " & vbCrLf &
           "		   AND n.EndEmpresa_Id   = nXi.EndEmpresa_Id                                            " & vbCrLf &
           "		   AND n.Cliente_Id      = nXi.Cliente_Id                                               " & vbCrLf &
           "		   AND n.EndCliente_Id   = nXi.EndCliente_Id                                            " & vbCrLf &
           "		   AND n.EntradaSaida_Id = nXi.EntradaSaida_Id                                          " & vbCrLf &
           "		   AND n.Serie_Id        = nXi.Serie_Id                                                 " & vbCrLf &
           "		   AND n.Nota_Id         = nXi.Nota_Id                                                  " & vbCrLf &
           "		 INNER JOIN NotasFiscaisXEncargos nXe                                                   " & vbCrLf &
           "		 	ON nXe.Empresa_Id      = nXi.Empresa_Id                                             " & vbCrLf &
           "		   AND nXe.EndEmpresa_Id   = nXi.EndEmpresa_Id                                          " & vbCrLf &
           "		   AND nXe.Cliente_Id      = nXi.Cliente_Id                                             " & vbCrLf &
           "		   AND nXe.EndCliente_Id   = nXi.EndCliente_Id                                          " & vbCrLf &
           "		   AND nXe.EntradaSaida_Id = nXi.EntradaSaida_Id                                        " & vbCrLf &
           "		   AND nXe.Serie_Id        = nXi.Serie_Id                                               " & vbCrLf &
           "		   AND nXe.Nota_Id         = nXi.Nota_Id                                                " & vbCrLf &
           "		   AND nXe.Produto_Id      = nXi.Produto_Id                                             " & vbCrLf &
           "		   AND nXe.CFOP_Id         = nXi.CFOP_Id                                                " & vbCrLf &
           "		   And nXe.sequencia_ID    = nXi.sequencia_ID                                           " & vbCrLf &
           "		 INNER JOIN SubOperacoes sOp                                                            " & vbCrLf &
           "			ON sOp.Operacao_Id     = n.Operacao                                                 " & vbCrLf &
           "		   AND sOp.SubOperacoes_Id = n.SubOperacao                                              " & vbCrLf &
           "		 INNER JOIN #Posicao                                                                    " & vbCrLf &
           "		    ON #Posicao.Empresa_Id    = n.Empresa_Id                                            " & vbCrLf &
           "		   AND #Posicao.EndEmpresa_Id = n.endEmpresa_Id                                         " & vbCrLf &
           "		   AND #Posicao.Pedido_Id     = n.Pedido                                                " & vbCrLf &
           "		 WHERE (n.Situacao = 1)                                                                 " & vbCrLf &
           "		   AND nXe.Encargo_Id = 'LIQUIDO'                                                       " & vbCrLf &
           "		   And (sOp.Classe NOT IN ('" & eClassesOperacoes.GLOBAL.ToString & "', '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "', '" & eClassesOperacoes.CONTAEORDEM.ToString & "'))  " & vbCrLf &
           "		 GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Pedido                                       " & vbCrLf &
           "                                                                                               " & vbCrLf &
           "		)as consulta                                                                            " & vbCrLf &
           "  INNER JOIN #Posicao                                                                          " & vbCrLf &
           "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id                                            " & vbCrLf &
           "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id                                         " & vbCrLf &
           "	AND	consulta.Pedido        = #Posicao.Pedido_Id                                             " & vbCrLf

        'Atualiza Quantidade Entregue Romaneios ----
        'PARTE 6 - LINHA 935
        Sql &= "Update #Posicao SET                                      " & vbCrLf &
                "	#Posicao.Laudo   = consulta.Laudo                     " & vbCrLf &
                "  from (SELECT n.Empresa_Id,                             " & vbCrLf &
                "               n.EndEmpresa_Id,                          " & vbCrLf &
                "               n.Pedido,                                 " & vbCrLf &
                "			   SUM(CASE                                   " & vbCrLf &
                "			         WHEN sOp.Devolucao = 'S'             " & vbCrLf &
                "			           THEN r.PesoLiquido * - 1           " & vbCrLf &
                "			           ELSE r.PesoLiquido                 " & vbCrLf &
                "			       END) AS Laudo                          " & vbCrLf &
                "		  FROM NotasFiscais n                             " & vbCrLf &
                "		 INNER JOIN SubOperacoes sOp                      " & vbCrLf &
                "			ON sOp.Operacao_Id     = n.Operacao           " & vbCrLf &
                "		   AND sOp.SubOperacoes_Id = n.SubOperacao        " & vbCrLf &
                "		 INNER JOIN NotasFiscaisXRomaneios nXr            " & vbCrLf &
                "			ON nXr.Empresa_Id      = n.Empresa_Id         " & vbCrLf &
                "		   AND nXr.EndEmpresa_Id   = n.EndEmpresa_Id      " & vbCrLf &
                "		   AND nXr.Cliente_Id      = n.Cliente_Id         " & vbCrLf &
                "		   AND nXr.EndCliente_Id   = n.EndCliente_Id      " & vbCrLf &
                "		   AND nXr.EntradaSaida_Id = n.EntradaSaida_Id    " & vbCrLf &
                "		   AND nXr.Serie_Id        = n.Serie_Id           " & vbCrLf &
                "		   AND nXr.Nota_Id         = n.Nota_Id            " & vbCrLf &
                "		 INNER JOIN Romaneios r                           " & vbCrLf &
                "			ON r.Empresa_Id    = nXr.Empresa_Id           " & vbCrLf &
                "		   AND r.EndEmpresa_Id = nXr.EndEmpresa_Id        " & vbCrLf &
                "		   AND r.Romaneio_Id   = nXr.Romaneio_Id          " & vbCrLf &
                "		 INNER JOIN #Posicao                              " & vbCrLf &
                "		    ON #Posicao.Empresa_Id    = n.Empresa_Id      " & vbCrLf &
                "		   AND #Posicao.EndEmpresa_Id = n.endEmpresa_Id   " & vbCrLf &
                "		   AND #Posicao.Pedido_Id     = n.Pedido          " & vbCrLf &
                "		 Where (n.Situacao = 1)                           " & vbCrLf &
                "		 GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Pedido " & vbCrLf &
                "		) consulta                                        " & vbCrLf &
                "  INNER JOIN #Posicao                                    " & vbCrLf &
                "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id      " & vbCrLf &
                "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id   " & vbCrLf &
                "	AND	consulta.Pedido        = #Posicao.Pedido_Id       " & vbCrLf

        'Ajustando Saldos ----
        'PARTE 7 - LINHA 960
        Sql &= "UPDATE #Posicao SET                              " & vbCrLf &
                "   #Posicao.Contratado = #Posicao.Entregue       " & vbCrLf &
                " where #Posicao.Contratado < #Posicao.Entregue   " & vbCrLf

        Sql &= "UPDATE #Posicao SET                                                                              " & vbCrLf &
                "	#Posicao.AEntregar = #Posicao.Contratado -                                                    " & vbCrLf &
                "												Case                                              " & vbCrLf &
                "												  when #Posicao.EntradaSaida = 'E'                " & vbCrLf &
                "													then Case                                     " & vbCrLf &
                "														   when p.FreteCifFob = 'FOB'             " & vbCrLf &
                "			  												 then ISNULL(#Posicao.Entregue, 0)    " & vbCrLf &
                "															  else ISNULL(#Posicao.Laudo, 0)      " & vbCrLf &
                "														 End                                      " & vbCrLf &
                "													else Case                                     " & vbCrLf &
                "														   when p.FreteCifFob = 'FOB'             " & vbCrLf &
                "															 then ISNULL(#Posicao.Entregue, 0)    " & vbCrLf &
                "															 else ISNULL(#Posicao.Entregue, 0)    " & vbCrLf &
                "														 End                                      " & vbCrLf &
                "												End                                               " & vbCrLf &
                "	FROM #Posicao                                                                                 " & vbCrLf &
                "   INNER JOIN Pedidos p                                                                          " & vbCrLf &
                " 	  ON p.Empresa_Id    = #Posicao.Empresa_Id                                                    " & vbCrLf &
                "	 AND p.EndEmpresa_Id = #Posicao.EndEmpresa_Id                                                 " & vbCrLf &
                "	 AND p.Pedido_Id     = #Posicao.Pedido_Id                                                     " & vbCrLf

        Sql &= "UPDATE #Posicao SET                                                               " & vbCrLf &
            "   #Posicao.AFixar = #Posicao.Contratado - isnull(#Posicao.QuantidadeFixado, 0)   " & vbCrLf &
            "                                                                                  " & vbCrLf &
            "Delete #Posicao                                                                   " & vbCrLf &
            " where #Posicao.Contratado = 0                                                    " & vbCrLf &
            "   and isnull(#Posicao.Entregue, 0) = 0                                           " & vbCrLf

        'Atualiza Contas a Receber ----
        'PARTE 8 - LINHA 993
        Sql &= "UPDATE #Posicao SET                                                    " & vbCrLf &
                 "	#Posicao.Programado_ReceberOficial  = consulta.ProgramadoOficial,   " & vbCrLf &
                 "	#Posicao.Baixado_ReceberOficial     = consulta.BaixadoOficial,      " & vbCrLf &
                 "	#Posicao.Programado_ReceberMoeda    = consulta.ProgramadoMoeda,     " & vbCrLf &
                 "	#Posicao.Baixado_ReceberMoeda       = consulta.BaixadoMoeda         " & vbCrLf &
                 " FROM (SELECT cR.EmpresaPedido,                                        " & vbCrLf &
                 "              cR.EndEmpresaPedido,                                     " & vbCrLf &
                 "              cR.Pedido,                                               " & vbCrLf &
                 "			  ISNULL(SUM(CASE                                           " & vbCrLf &
                 "			               WHEN cR.Provisao <> 1                        " & vbCrLf &
                 "			                 THEN cR.ValorDoDocumento                   " & vbCrLf &
                 "			                 ELSE 0                                     " & vbCrLf &
                 "			             END), 0) AS ProgramadoOficial,                 " & vbCrLf &
                 "			  ISNULL(SUM(CASE                                           " & vbCrLf &
                 "			               WHEN cR.Provisao = 1                         " & vbCrLf &
                 "			                 THEN cR.ValorDoDocumento                   " & vbCrLf &
                 "			                 ELSE 0                                     " & vbCrLf &
                 "			             END), 0) AS BaixadoOficial,                    " & vbCrLf &
                 "			  ISNULL(SUM(CASE                                           " & vbCrLf &
                 "			               WHEN cR.Provisao <> 1                        " & vbCrLf &
                 "			                 THEN cR.MoedaValorDoDocumento              " & vbCrLf &
                 "			                 ELSE 0                                     " & vbCrLf &
                 "			             END), 0) AS ProgramadoMoeda,                   " & vbCrLf &
                 "			  ISNULL(SUM(CASE                                           " & vbCrLf &
                 "			               WHEN cR.Provisao = 1                         " & vbCrLf &
                 "			                 THEN cR.MoedaValorDoDocumento              " & vbCrLf &
                 "			                 ELSE 0                                     " & vbCrLf &
                 "			              END), 0) AS BaixadoMoeda                      " & vbCrLf &
                 "		 FROM ContasAReceber cR                                         " & vbCrLf &
                 "		INNER JOIN #Posicao                                             " & vbCrLf &
                 "		   ON cR.EmpresaPedido    = #Posicao.Empresa_Id                 " & vbCrLf &
                 "		  AND cR.EndEmpresaPedido = #Posicao.EndEmpresa_Id              " & vbCrLf &
                 "		  AND cR.Pedido           = #Posicao.Pedido_Id                  " & vbCrLf &
                 "		WHERE cR.Situacao = 1                                           " & vbCrLf &
                 "		GROUP BY cR.EmpresaPedido, cR.EndEmpresaPedido, cR.Pedido       " & vbCrLf &
                 "	   ) AS consulta                                                    " & vbCrLf &
                 " INNER JOIN #Posicao                                                   " & vbCrLf &
                 "	ON consulta.EmpresaPedido    = #Posicao.Empresa_Id                  " & vbCrLf &
                 "   AND consulta.EndEmpresaPedido = #Posicao.EndEmpresa_Id              " & vbCrLf &
                 "   AND consulta.Pedido           = #Posicao.Pedido_Id                  " & vbCrLf

        'Atualiza Contas a Pagar ----
        'PARTE 9 - LINHA 1021
        Sql &= "UPDATE #Posicao SET                                                 " & vbCrLf &
              "	#Posicao.Programado_PagarOficial = consulta.ProgramadoOficial,   " & vbCrLf &
              "	#Posicao.Baixado_PagarOficial    = consulta.BaixadoOficial,      " & vbCrLf &
              "	#Posicao.Programado_PagarMoeda   = consulta.ProgramadoMoeda,     " & vbCrLf &
              "	#Posicao.Baixado_PagarMoeda      = consulta.BaixadoMoeda         " & vbCrLf &
              "  FROM (SELECT cP.EmpresaPedido,                                    " & vbCrLf &
              "               cP.EndEmpresaPedido,                                 " & vbCrLf &
              "               cP.Pedido,                                           " & vbCrLf &
              "			   ISNULL(SUM(CASE                                       " & vbCrLf &
              "			                WHEN cP.Provisao <> 1                    " & vbCrLf &
              "			                  THEN cP.ValorDoDocumento               " & vbCrLf &
              "			                  ELSE 0                                 " & vbCrLf &
              "			              END), 0) AS ProgramadoOficial,             " & vbCrLf &
              "			   ISNULL(SUM(CASE                                       " & vbCrLf &
              "			                WHEN cP.Provisao = 1                     " & vbCrLf &
              "			                  THEN cP.ValorDoDocumento               " & vbCrLf &
              "			                  ELSE 0                                 " & vbCrLf &
              "			              END), 0) AS BaixadoOficial,                " & vbCrLf &
              "			   ISNULL(SUM(CASE                                       " & vbCrLf &
              "			                WHEN cP.Provisao <> 1                    " & vbCrLf &
              "			                  THEN cP.MoedaValorDoDocumento          " & vbCrLf &
              "			                  ELSE 0                                 " & vbCrLf &
              "			              END), 0) AS ProgramadoMoeda,               " & vbCrLf &
              "			   ISNULL(SUM(CASE                                       " & vbCrLf &
              "			                WHEN cP.Provisao = 1                     " & vbCrLf &
              "			                  THEN cP.MoedaValorDoDocumento          " & vbCrLf &
              "			                  ELSE 0                                 " & vbCrLf &
              "			              END), 0) AS BaixadoMoeda                   " & vbCrLf &
              "		  FROM ContasAPagar cP                                       " & vbCrLf &
              "	     INNER JOIN #Posicao                                         " & vbCrLf &
              "		    ON cP.EmpresaPedido    = #Posicao.Empresa_Id             " & vbCrLf &
              "		   AND cP.EndEmpresaPedido = #Posicao.EndEmpresa_Id          " & vbCrLf &
              "		   AND cP.Pedido           = #Posicao.Pedido_Id              " & vbCrLf &
              "		 WHERE cP.Situacao = 1                                       " & vbCrLf &
              "		   And cP.Carteira not in ('001001004')                      " & vbCrLf &
              "		 GROUP BY cP.EmpresaPedido, cP.EndEmpresaPedido, cP.Pedido   " & vbCrLf &
              "		) AS consulta                                                " & vbCrLf &
              "  INNER JOIN #Posicao                                               " & vbCrLf &
              "	 ON consulta.EmpresaPedido    = #Posicao.Empresa_Id              " & vbCrLf &
              "	AND consulta.EndEmpresaPedido = #Posicao.EndEmpresa_Id           " & vbCrLf &
              "	AND	consulta.Pedido           = #Posicao.Pedido_Id               " & vbCrLf

        'Atualiza Adiantamentos de Materia Prima ----
        'PARTE 10 - LINHA 1050
        Sql &= "UPDATE #Posicao SET                                               " & vbCrLf &
              "	#Posicao.Adiantamento = consulta.Valor                         " & vbCrLf &
              "  FROM (SELECT cP.EmpresaPedido,                                  " & vbCrLf &
              "               cP.EndEmpresaPedido,                               " & vbCrLf &
              "               cP.Pedido as Pedido,                               " & vbCrLf &
              "               Sum(DebitoOficial - CreditoOficial) as Valor       " & vbCrLf &
              "		  FROM ContasAPagar cP                                     " & vbCrLf &
              "		 INNER JOIN Razao r                                        " & vbCrLf &
              "			ON r.Titulo = cP.Registro_Id                           " & vbCrLf &
              "		 WHERE Conta_ID  = '1010303'                               " & vbCrLf &
              "		 Group by cP.EmpresaPedido, cP.EndEmpresaPedido, cp.Pedido " & vbCrLf &
              "		) AS consulta                                              " & vbCrLf &
              "  INNER JOIN #Posicao                                             " & vbCrLf &
              "	 ON consulta.EmpresaPedido    = #Posicao.Empresa_Id            " & vbCrLf &
              "	AND consulta.EndEmpresaPedido = #Posicao.EndEmpresa_Id         " & vbCrLf &
              "	AND	consulta.Pedido           = #Posicao.Pedido_Id             " & vbCrLf

        'Ajusta Saldo Financeiro ----------------
        'PARTE 11 - LINHA 1072
        Sql &= "UPDATE #Posicao                                                                                                                        " & vbCrLf &
            "	SET #Posicao.Programado_SaldoOficial = isnull(#Posicao.Programado_PagarOficial, 0) - isnull(#Posicao.Programado_ReceberOficial, 0), " & vbCrLf &
            "		#Posicao.Programado_SaldoMoeda   = isnull(#Posicao.Programado_PagarMoeda, 0)   - isnull(#Posicao.Programado_ReceberMoeda, 0),   " & vbCrLf &
            "		#Posicao.Baixado_SaldoOficial    = isnull(#Posicao.Baixado_PagarOficial, 0)    - isnull(#Posicao.Baixado_ReceberOficial, 0),    " & vbCrLf &
            "		#Posicao.Baixado_SaldoMoeda      = isnull(#Posicao.Baixado_PagarMoeda, 0)      - isnull(#Posicao.Baixado_ReceberMoeda, 0)       " & vbCrLf &
            "from #Posicao                                                                                                                        " & vbCrLf

        Sql &= "Update #Posicao Set                                                           " & vbCrLf &
           "  #Posicao.Programado_SaldoOficial = (#Posicao.Programado_SaldoOficial * -1)  " & vbCrLf &
           "  from #Posicao                                                               " & vbCrLf &
           " where #Posicao.Programado_SaldoOficial < 0                                   " & vbCrLf

        Sql &= "Update #Posicao Set                                                         " & vbCrLf &
             "  #Posicao.Programado_SaldoMoeda   = (#Posicao.Programado_SaldoMoeda * -1)  " & vbCrLf &
             "  from #Posicao                                                             " & vbCrLf &
             " where #Posicao.Programado_SaldoMoeda < 0                                   " & vbCrLf

        Sql &= "Update #Posicao Set                                                        " & vbCrLf &
            "  #Posicao.Baixado_SaldoOficial = (#Posicao.Baixado_SaldoOficial * -1)     " & vbCrLf &
            "  from #Posicao                                                            " & vbCrLf &
            " where #Posicao.Baixado_SaldoOficial < 0                                   " & vbCrLf

        Sql &= "Update #Posicao Set                                                 " & vbCrLf &
           "  #Posicao.Baixado_SaldoMoeda   = (#Posicao.Baixado_SaldoMoeda * -1)" & vbCrLf &
           "  from #Posicao                                                     " & vbCrLf &
           " where #Posicao.Baixado_SaldoMoeda < 0                              " & vbCrLf

        'Calcula Quantidade Paga ----------------
        'PARTE 12 - LINHA 1094
        Sql &= "UPDATE #Posicao SET                                                                                                   " & vbCrLf &
           "	#Posicao.Pago = CASE                                                                                               " & vbCrLf &
           "					  WHEN #Posicao.MoedaPedido = 1                                                                    " & vbCrLf &
           "						THEN (#Posicao.QuantidadeFixado / #Posicao.ValorFixadoOficial) * #Posicao.Baixado_SaldoOficial " & vbCrLf &
           "						ELSE (#Posicao.QuantidadeFixado / #Posicao.ValorFixadoMoeda)   * #Posicao.Baixado_SaldoMoeda   " & vbCrLf &
           "					END                                                                                                " & vbCrLf &
           "WHERE #Posicao.ValorFixadoOficial > 0                                                                                 " & vbCrLf

        Sql &= "UPDATE #Posicao Set                                                " & vbCrLf &
          "  #Posicao.RecebidoNaoPago = (#Posicao.Entregue - #Posicao.Pago)   " & vbCrLf &
          " WHERE isnull(#Posicao.Entregue, 0) > #Posicao.Pago                " & vbCrLf

        SqlArray.Add(Sql)

        Sql &= "UPDATE #Posicao Set                                                          " & vbCrLf &
            " #Posicao.PagoNaoRecebido = (#Posicao.Pago - isnull(#Posicao.Entregue, 0))   " & vbCrLf &
            " WHERE #Posicao.Pago > isnull(#Posicao.Entregue, 0)                          " & vbCrLf

        'PARTE 13 - LINHA 1109
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
                "   #Posicao.TotalOficial = (SELECT SUM(CASE                                     " & vbCrLf &
                "                                         WHEN pXi.TipoDeLancamento = 'E'        " & vbCrLf &
                "                                           THEN (pXi.TotalOficial * - 1)        " & vbCrLf &
                "                                           ELSE pXi.TotalOficial                " & vbCrLf &
                "                                       END) AS Oficial                          " & vbCrLf &
                "							  FROM PedidoXItemxLancamento pXi                             " & vbCrLf &
                "							 WHERE (pXi.Empresa_Id    = #Posicao.Empresa_Id)     " & vbCrLf &
                "							   AND (pXi.EndEmpresa_Id = #Posicao.EndEmpresa_Id)  " & vbCrLf &
                "							   AND (pXi.Pedido_Id     = #Posicao.Pedido_Id)),    " & vbCrLf &
                "	#Posicao.TotalMoeda  = (SELECT SUM(CASE                                      " & vbCrLf &
                "	                                     WHEN pXi.TipoDeLancamento = 'E'         " & vbCrLf &
                "	                                       THEN (pXi.TotalMoeda * - 1)           " & vbCrLf &
                "	                                       ELSE TotalMoeda                       " & vbCrLf &
                "	                                    END) AS Moeda                            " & vbCrLf &
                "							  FROM PedidoXItemxLancamento pXi                             " & vbCrLf &
                "							 WHERE (pXi.Empresa_Id = #Posicao.Empresa_Id)        " & vbCrLf &
                "							   AND (pXi.EndEmpresa_Id = #Posicao.EndEmpresa_Id)  " & vbCrLf &
                "							   AND (pXi.Pedido_Id = #Posicao.Pedido_Id))         " & vbCrLf

        Sql &= "UPDATE #Posicao SET                                                                            " & vbCrLf &
           "   #Posicao.UnitarioOficial = (#Posicao.TotalOficial / #Posicao.QuantidadeFixado) " & vbCrLf &
           "  FROM #Posicao                                                                                    " & vbCrLf &
           " INNER JOIN Produtos p                                                                             " & vbCrLf &
           "    ON p.Produto_Id = #Posicao.Produto                                                             " & vbCrLf &
           " Where #Posicao.TotalOficial > 0                                                                   " & vbCrLf &
           "   and #Posicao.QuantidadeFixado   > 0                                                             " & vbCrLf

        Sql &= "UPDATE #Posicao SET                                                                          " & vbCrLf &
                "  #Posicao.UnitarioMoeda = (#Posicao.TotalMoeda / #Posicao.Contratado)      " & vbCrLf &
                "  FROM #Posicao                                                                              " & vbCrLf &
                " INNER JOIN Produtos p                                                                       " & vbCrLf &
                "	ON p.Produto_Id = #Posicao.Produto                                                        " & vbCrLf &
                " Where #Posicao.TotalMoeda > 0                                                               " & vbCrLf &
                "   and #Posicao.Contratado > 0                                                               " & vbCrLf
    End Sub

    Protected Sub lnkPosição_Click(sender As Object, e As EventArgs) Handles lnkPosição.Click
        Try
            If Not ValidaCampos() Then
                Exit Sub
            End If

            Dim strEmpresa() As String = ddlEmpresa.SelectedValue().Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            'AtualizaPosicaoDeComissoes()

            Dim sqlSelect = ""

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sqlSelect &= " AND Left(COM.Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'"
                Else
                    sqlSelect &= " AND COM.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                 " AND COM.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1)
                End If
            End If

            If Not chkTodosRepresentantes.Checked Then
                If Not String.IsNullOrWhiteSpace(ddlRepresentante.SelectedValue) Then
                    sqlSelect &= " AND COM.Representante_Id = '" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                 " AND COM.EndRepresentante_Id = " & ddlRepresentante.SelectedValue.Split("-")(1)
                End If
            End If

            If Not (String.IsNullOrWhiteSpace(txtDataInicial.Text) And String.IsNullOrWhiteSpace(txtDataFinal.Text)) Then
                sqlSelect &= " AND PED.DataPedido BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                             " AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'"
            End If

            Dim sql = "SELECT " & vbCrLf &
                  "    EMP.Reduzido + '-' + EMP.Estado AS Empresa, " & vbCrLf &
                  "    REP.Cliente_Id AS Representante, " & vbCrLf &
                  "    REP.Nome, " & vbCrLf &
                  "    REP.Cidade, " & vbCrLf &
                  "    REP.Estado AS UF, " & vbCrLf &
                  "    PED.PedidoEfetivo AS Pedido, " & vbCrLf &
                  "    PED.DataPedido, " & vbCrLf &
                  "    CONVERT(INT, PIL.Contratado) AS Contratado, " & vbCrLf &
                  "    CONVERT(INT, NFXI.Entregue) AS Entregue, " & vbCrLf &
                  "    COM.ValorComissao AS ComXTon, " & vbCrLf &
                  "    ((NFXI.Entregue * COM.ValorComissao) / 1000) AS Comissao, " & vbCrLf &
                  "    CASE WHEN NFXT.ValorPago > 0 THEN (((NFXI.Entregue * COM.ValorComissao) / 1000) * ((NFXT.ValorPago / NFXT.ValorTotal) * 100)) / 100 ELSE 0 END AS Liberado, " & vbCrLf &
                  "    CASE WHEN NFXT.ValorPago > 0 THEN ((NFXI.Entregue * COM.ValorComissao) / 1000) - (((NFXI.Entregue * COM.ValorComissao) / 1000) * ((NFXT.ValorPago / NFXT.ValorTotal) * 100)) / 100 ELSE 0 END AS ALiberar, " & vbCrLf &
                  "    0 AS Faturada, " & vbCrLf &
                  "    0 AS ComPaga, " & vbCrLf &
                  "    NFXT.ValorTotal, " & vbCrLf &
                  "    NFXT.ValorPago, " & vbCrLf &
                  "    CASE WHEN NFXT.ValorPago > 0 THEN ((NFXT.ValorPago / NFXT.ValorTotal) * 100) ELSE 0 END AS PerPago " & vbCrLf &
                  "FROM " & vbCrLf &
                  "    Comissoes COM " & vbCrLf &
                  "    INNER JOIN Pedidos AS PED ON " & vbCrLf &
                  "        PED.Empresa_Id = COM.Empresa_Id AND " & vbCrLf &
                  "        PED.EndEmpresa_Id = COM.EndRepresentante_Id AND " & vbCrLf &
                  "        PED.Pedido_Id = COM.Pedido_Id " & vbCrLf &
                  "    INNER JOIN PedidoXItem PXI ON " & vbCrLf &
                  "        PED.Empresa_Id = PXI.Empresa_Id AND " & vbCrLf &
                  "        PED.EndEmpresa_Id = PXI.EndEmpresa_Id AND " & vbCrLf &
                  "        PED.Pedido_Id = PXI.Pedido_Id " & vbCrLf &
                  "    INNER JOIN SubOperacoes SOP ON " & vbCrLf &
                  "        SOP.Operacao_Id = PED.Operacao AND " & vbCrLf &
                  "        SOP.SubOperacoes_Id = PED.SubOperacao " & vbCrLf &
                  "    OUTER APPLY ( " & vbCrLf &
                  "        SELECT " & vbCrLf &
                  "            Sum(CASE WHEN PIL.TipoDeLancamento = 'E' THEN Quantidade * -1 ELSE Quantidade END) AS Contratado " & vbCrLf &
                  "        FROM " & vbCrLf &
                  "            PedidoXItemXLancamento PIL " & vbCrLf &
                  "        WHERE " & vbCrLf &
                  "            PXI.Empresa_Id = PIL.Empresa_Id AND " & vbCrLf &
                  "            PXI.EndEmpresa_Id = PIL.EndEmpresa_Id AND " & vbCrLf &
                  "            PXI.Pedido_Id = PIL.Pedido_Id AND " & vbCrLf &
                  "            PXI.Produto_ID = PIL.Produto_Id " & vbCrLf &
                  "    ) AS PIL " & vbCrLf &
                  "    OUTER APPLY ( " & vbCrLf &
                  "        SELECT " & vbCrLf &
                  "            isnull(Sum(CASE WHEN NFI.EntradaSaida_Id = 'S' THEN NFI.PesoFiscal ELSE NFI.PesoFiscal * -1 END), 0) AS Entregue " & vbCrLf &
                  "        FROM " & vbCrLf &
                  "            NotasFiscais NF " & vbCrLf &
                  "            INNER JOIN RTGraos.dbo.NotasFiscaisXItens NFI ON " & vbCrLf &
                  "                NF.Empresa_Id = NFI.Empresa_Id AND " & vbCrLf &
                  "                NF.EndEmpresa_Id = NFI.EndEmpresa_Id AND " & vbCrLf &
                  "                NF.Cliente_Id = NFI.Cliente_Id AND " & vbCrLf &
                  "                NF.EndCliente_Id = NFI.EndCliente_Id AND " & vbCrLf &
                  "                NF.EntradaSaida_Id = NFI.EntradaSaida_Id AND " & vbCrLf &
                  "                NF.Serie_Id = NFI.Serie_Id AND " & vbCrLf &
                  "                NF.Nota_Id = NFI.Nota_Id AND " & vbCrLf &
                  "                NF.TipoDeDocumento <> 15 " & vbCrLf &
                  "        WHERE " & vbCrLf &
                  "            PED.Empresa_Id = NF.Empresa_Id AND " & vbCrLf &
                  "            PED.EndEmpresa_Id = NF.EndEmpresa_Id AND " & vbCrLf &
                  "            PED.Cliente = NF.Cliente_Id AND " & vbCrLf &
                  "            PED.EndCliente = NF.EndCliente_Id AND " & vbCrLf &
                  "            PED.Pedido_Id = NF.Pedido AND " & vbCrLf &
                  "            NFI.Produto_Id = PXI.Produto_ID AND " & vbCrLf &
                  "            NF.TipoDeDocumento <> 15 " & vbCrLf &
                  "    ) AS NFXI " & vbCrLf &
                  "    OUTER APPLY ( " & vbCrLf &
                  "        SELECT " & vbCrLf &
                  "            isnull(Sum(CASE WHEN CAR.Provisao IN (1, 2) THEN Car.ValorDoDocumento ELSE 0 END), 0) AS ValorTotal, " & vbCrLf &
                  "            isnull(Sum(CASE WHEN CAR.Provisao IN (1) THEN Car.ValorLiquido ELSE 0 END), 0) AS ValorPago " & vbCrLf &
                  "        FROM " & vbCrLf &
                  "            NotasFiscais NF " & vbCrLf &
                  "            INNER JOIN NotaFiscalXTitulo NFXT ON " & vbCrLf &
                  "                NF.Empresa_Id = NFXT.Empresa_Id AND " & vbCrLf &
                  "                NF.EndEmpresa_Id = NFXT.EndEmpresa_Id AND " & vbCrLf &
                  "                NF.Cliente_Id = NFXT.Cliente_Id AND " & vbCrLf &
                  "                NF.EndCliente_Id = NFXT.EndCliente_Id AND " & vbCrLf &
                  "                NF.EntradaSaida_Id = NFXT.EntradaSaida_Id AND " & vbCrLf &
                  "                NF.Serie_Id = NFXT.Serie_Id AND " & vbCrLf &
                  "                NF.Nota_Id = NFXT.Nota_Id " & vbCrLf &
                  "            INNER JOIN ContasAReceber CAR ON " & vbCrLf &
                  "                NFXT.Titulo_Id = CAR.Registro_Id AND " & vbCrLf &
                  "                CAR.Situacao = 1 " & vbCrLf &
                  "        WHERE " & vbCrLf &
                  "            PED.Empresa_Id = NF.Empresa_Id AND " & vbCrLf &
                  "            PED.EndEmpresa_Id = NF.EndEmpresa_Id AND " & vbCrLf &
                  "            PED.Cliente = NF.Cliente_Id AND " & vbCrLf &
                  "            PED.EndCliente = NF.EndCliente_Id AND " & vbCrLf &
                  "            PED.Pedido_Id = NF.Pedido " & vbCrLf &
                  "    ) AS NFXT " & vbCrLf &
                  "    OUTER APPLY ( " & vbCrLf &
                  "        SELECT " & vbCrLf &
                  "            isnull(Sum(CASE WHEN CAP.Provisao IN (1, 2) THEN CAP.ValorDoDocumento ELSE 0 END), 0) AS ComProgramada, " & vbCrLf &
                  "            isnull(Sum(CASE WHEN CAP.Provisao IN (1) THEN CAP.ValorLiquido ELSE 0 END), 0) AS ComPaga " & vbCrLf &
                  "        FROM " & vbCrLf &
                  "            NotasFiscais NF " & vbCrLf &
                  "            INNER JOIN NotaFiscalXTitulo NFXT ON " & vbCrLf &
                  "                NF.Empresa_Id = NFXT.Empresa_Id AND " & vbCrLf &
                  "                NF.EndEmpresa_Id = NFXT.EndEmpresa_Id AND " & vbCrLf &
                  "                NF.Cliente_Id = COM.Representante_Id AND " & vbCrLf &
                  "                NF.EndCliente_Id = 0 AND " & vbCrLf &
                  "                NF.EntradaSaida_Id = NFXT.EntradaSaida_Id AND " & vbCrLf &
                  "                NF.Serie_Id = NFXT.Serie_Id AND " & vbCrLf &
                  "                NF.Nota_Id = NFXT.Nota_Id " & vbCrLf &
                  "            INNER JOIN ContasAPagar CAP ON " & vbCrLf &
                  "                NFXT.Titulo_Id = CAP.Registro_Id AND " & vbCrLf &
                  "                CAP.Situacao = 1 " & vbCrLf &
                  "    ) AS NFXCAP " & vbCrLf &
                  "    INNER JOIN Clientes EMP ON " & vbCrLf &
                  "        PED.Empresa_Id = EMP.Cliente_Id AND " & vbCrLf &
                  "        PED.EndEmpresa_Id = EMP.Endereco_Id " & vbCrLf &
                  "    INNER JOIN Clientes REP ON " & vbCrLf &
                  "        COM.Representante_Id = REP.Cliente_Id AND " & vbCrLf &
                  "        COM.EndRepresentante_Id = REP.Endereco_Id " & vbCrLf &
                  "WHERE " & vbCrLf &
                  "    PIL.Contratado > 0 " & vbCrLf &
                  "AND COM.ValorComissao > 0 " & vbCrLf &
                  "AND PED.Situacao = 1 " & vbCrLf &
                  sqlSelect & vbCrLf &
                  "UNION " & vbCrLf &
                  "SELECT " & vbCrLf &
                  "    EMP.Reduzido + '-' + EMP.Estado AS Empresa, " & vbCrLf &
                  "    REP.Cliente_Id AS Representante, " & vbCrLf &
                  "    REP.Nome, " & vbCrLf &
                  "    REP.Cidade, " & vbCrLf &
                  "    REP.Estado AS UF, " & vbCrLf &
                  "    'ZZZ-' + '-' + Convert(Varchar, NFXCAP.Pedido) AS Pedido, " & vbCrLf &
                  "    NFXCAP.Movimento AS DataPedido, " & vbCrLf &
                  "    0 AS Contratado, " & vbCrLf &
                  "    0 AS Entregue, " & vbCrLf &
                  "    0 AS ComXTon, " & vbCrLf &
                  "    0 AS Comissao, " & vbCrLf &
                  "    0 AS Liberado, " & vbCrLf &
                  "    0 AS ALiberar, " & vbCrLf &
                  "    Sum(NFXCAP.ComProgramada) AS Faturada, " & vbCrLf &
                  "    Sum(NFXCAP.ComPaga) AS ComPaga, " & vbCrLf &
                  "    0 AS ValorTotal, " & vbCrLf &
                  "    0 AS ValorPago, " & vbCrLf &
                  "    0 AS PerPago " & vbCrLf &
                  "FROM " & vbCrLf &
                  "    Comissoes COM " & vbCrLf &
                  "    OUTER APPLY ( " & vbCrLf &
                  "        SELECT " & vbCrLf &
                  "            isnull(Sum(CASE WHEN CAP.Provisao IN (1, 2) THEN CAP.ValorDoDocumento ELSE 0 END), 0) AS ComProgramada, " & vbCrLf &
                  "            isnull(Sum(CASE WHEN CAP.Provisao IN (1) THEN CAP.ValorLiquido ELSE 0 END), 0) AS ComPaga, " & vbCrLf &
                  "            NF.Pedido, " & vbCrLf &
                  "            NF.Movimento " & vbCrLf &
                  "        FROM " & vbCrLf &
                  "            NotasFiscais NF " & vbCrLf &
                  "            INNER JOIN NotaFiscalXTitulo NFXT ON " & vbCrLf &
                  "                NF.Empresa_Id = NFXT.Empresa_Id AND " & vbCrLf &
                  "                NF.EndEmpresa_Id = NFXT.EndEmpresa_Id AND " & vbCrLf &
                  "                NF.Cliente_Id = COM.Representante_Id AND " & vbCrLf &
                  "                NF.EndCliente_Id = 0 AND " & vbCrLf &
                  "                NF.EntradaSaida_Id = NFXT.EntradaSaida_Id AND " & vbCrLf &
                  "                NF.Serie_Id = NFXT.Serie_Id AND " & vbCrLf &
                  "                NF.Nota_Id = NFXT.Nota_Id " & vbCrLf &
                  "            INNER JOIN ContasAPagar CAP ON " & vbCrLf &
                  "                NFXT.Titulo_Id = CAP.Registro_Id AND " & vbCrLf &
                  "                CAP.Situacao = 1 " & vbCrLf &
                  "        WHERE " & vbCrLf &
                  "            NF.Movimento > '2025-01-01' " & vbCrLf &
                  "        GROUP BY " & vbCrLf &
                  "            NF.Pedido, " & vbCrLf &
                  "            NF.Movimento " & vbCrLf &
                  "    ) AS NFXCAP " & vbCrLf &
                  "    INNER JOIN Clientes EMP ON " & vbCrLf &
                  "        Com.Empresa_Id = EMP.Cliente_Id AND " & vbCrLf &
                  "        Com.EndEmpresa_Id = EMP.Endereco_Id " & vbCrLf &
                  "    INNER JOIN Clientes REP ON " & vbCrLf &
                  "        COM.Representante_Id = REP.Cliente_Id AND " & vbCrLf &
                  "        COM.EndRepresentante_Id = REP.Endereco_Id " & vbCrLf &
                  "    INNER JOIN Pedidos PED ON " & vbCrLf &
                  "        PED.Pedido_Id = NFXCAP.Pedido AND " & vbCrLf &
                  "        PED.Empresa_Id = COM.Empresa_Id AND" & vbCrLf &
                  "        PED.EndEmpresa_Id = COM.EndEmpresa_Id AND" & vbCrLf &
                  "        PED.Cliente = EMP.Cliente_Id " & vbCrLf &
                  "WHERE " & vbCrLf &
                  "    COM.ValorComissao > 0 " & vbCrLf &
                  sqlSelect & vbCrLf &
                  "GROUP BY " & vbCrLf &
                  "    EMP.Reduzido, " & vbCrLf &
                  "    EMP.Estado, " & vbCrLf &
                  "    REP.Cliente_Id, " & vbCrLf &
                  "    REP.Nome, " & vbCrLf &
                  "    REP.Cidade, " & vbCrLf &
                  "    REP.Estado, " & vbCrLf &
                  "    NFXCAP.Pedido, " & vbCrLf &
                  "    NFXCAP.Movimento " & vbCrLf &
                  "ORDER BY " & vbCrLf &
                  "    Nome, " & vbCrLf &
                  "    Pedido"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PosicaoDeComissao")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("POSIÇÃO DE COMISSÕES")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1} - {2} - {3}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo), objEmpresa.CodigoEndereco, objEmpresa.Reduzido)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO DE COMISSÕES")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " À " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:G" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    'criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            If col.ColumnName = "DataPedido" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0:dd/MM/yyyy}", CDate(row(col.ColumnName)))
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If
                            columnIndex += 1
                        Next

                        'aplicando formatação decimal nos campos de valores
                        worksheet.Cells(String.Format("J{0}:R{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'criando linha com totalizadores
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Formula = String.Format("=SUM(H6:H{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Formula = String.Format("=SUM(I6:I{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(J6:J{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Formula = String.Format("=SUM(K6:K{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Formula = String.Format("=SUM(L6:L{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Formula = String.Format("=SUM(M6:M{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Formula = String.Format("=SUM(N6:N{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Formula = String.Format("=SUM(O6:O{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Formula = String.Format("=SUM(P6:P{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("Q{0}", rowIndex)).Formula = String.Format("=SUM(Q6:Q{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("R{0}", rowIndex)).Formula = String.Format("=SUM(R6:R{0})", rowIndex - 1)

                    'formatando colunas dos totalizadores
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns()

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using

            'download do arquivo pelo browser
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub resumoExcel()
        Try
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue().Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    Sql = "-- 1 Busca os Titulos a Receber " & vbCrLf &
                    "select cp.Empresa, cp.EndEmpresa, cp.Cliente, cp.EndCliente, cp.Registro_id, cp.ValorDoDocumento, cp.ValorLiquido, cp.Historico, isnull(com.Percentual,0) As Percentual, " & vbCrLf &
                    "		nt.Empresa_Id, nt.EndEmpresa_Id, nt.Cliente_Id, nt.EndCliente_Id, nt.EntradaSaida_Id, nt.Serie_Id, nt.Nota_Id, " & vbCrLf &
                    "		ped.Pedido_Id, isnull(com.Representante_Id,'') AS Representante, isnull(com.EndRepresentante_id,0) AS EndRepresentante " & vbCrLf &
                    "into #temp " & vbCrLf &
                    "from contasAreceber cp " & vbCrLf &
                    "		inner join Pedidos ped " & vbCrLf &
                    "				on ped.Empresa_Id     = cp.EmpresaPedido " & vbCrLf &
                    "				and ped.EndEmpresa_Id = cp.EndEmpresaPedido " & vbCrLf &
                    "				and ped.Pedido_Id     = cp.Pedido	" & vbCrLf &
                    "		left join Comissoes com " & vbCrLf &
                    "				on com.Empresa_Id     = ped.Empresa_Id " & vbCrLf &
                    "				and com.EndEmpresa_Id = ped.EndEmpresa_Id " & vbCrLf &
                    "				and com.Pedido_Id     = ped.Pedido_Id " & vbCrLf &
                    "		left join notafiscalXtitulo nt " & vbCrLf &
                    "				on nt.Empresa_Id     = cp.Empresa " & vbCrLf &
                    "				and nt.EndEmpresa_Id = cp.EndEmpresa " & vbCrLf &
                    "				and nt.Titulo_Id     = cp.Registro_Id " & vbCrLf &
                    " Where (cp.Situacao = 1) " & vbCrLf &
                    "    AND (cp.Grupado <> 'M') " & vbCrLf &
                    "    AND (cp.Baixa BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
                    "    and cp.Provisao = 1 " & vbCrLf &
                    "    and cp.Empresa = '" & strEmpresa(0) & "' " & vbCrLf &
                    "    and cp.EndEmpresa = " & strEmpresa(1) & vbCrLf

                    If Not String.IsNullOrWhiteSpace(ddlRepresentante.SelectedValue) Then
                        Dim strRepresentante() As String = ddlRepresentante.SelectedValue().Split("-")
                        Sql &= "    and com.Representante_Id = '" & strRepresentante(0) & "'" & vbCrLf &
                            "    and com.EndRepresentante_Id = " & strRepresentante(1) & vbCrLf
                    End If

                    Sql &= "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "-- 2 Agruapa notas para Ver o valor Faturado" & vbCrLf &
                    "select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id " & vbCrLf &
                    " into #Notas" & vbCrLf &
                    "	from #temp" & vbCrLf &
                    "group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id " & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    " -- 3 soma das notas fiscais" & vbCrLf &
                    "select ni.Empresa_Id, ni.EndEmpresa_Id, ni.Cliente_Id, ni.EndCliente_Id, sum(ni.valor) AS ValorNF " & vbCrLf &
                    " into #ValorNotas " & vbCrLf &
                    "	from #Notas " & vbCrLf &
                    "	inner join NotasFiscaisXItens ni " & vbCrLf &
                    "			on ni.Empresa_id       = #Notas.Empresa_id " & vbCrLf &
                    "			and ni.EndEmpresa_id   = #Notas.EndEmpresa_Id " & vbCrLf &
                    "			and ni.Cliente_Id      = #Notas.Cliente_Id " & vbCrLf &
                    "			and ni.EndCliente_Id   = #Notas.EndCliente_Id " & vbCrLf &
                    "			and ni.EntradaSaida_id = #Notas.EntradaSaida_id " & vbCrLf &
                    "			and ni.Serie_id        = #Notas.Serie_Id " & vbCrLf &
                    "			and ni.Nota_Id         = #Notas.Nota_id " & vbCrLf &
                    "group by ni.Empresa_Id, ni.EndEmpresa_Id, ni.Cliente_Id, ni.EndCliente_Id " & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "-- 4 Cria resumo por nota" & vbCrLf &
                    "select Empresa, EndEmpresa, Cliente, EndCliente, Registro_id, ValorDoDocumento, ValorLiquido, Historico, " & vbCrLf &
                    "		Pedido_id AS Pedido, Percentual, Representante, EndRepresentante, Nota_Id, " & vbCrLf &
                    "		(select  sum(ni.valor) from NotasFiscaisXItens ni " & vbCrLf &
                    "								where ni.Empresa_id       = #temp.Empresa_id " & vbCrLf &
                    "								   and ni.EndEmpresa_id   = #temp.EndEmpresa_Id " & vbCrLf &
                    "								   and ni.Cliente_Id      = #temp.Cliente_Id " & vbCrLf &
                    "								   and ni.EndCliente_Id   = #temp.EndCliente_Id " & vbCrLf &
                    "								   and ni.EntradaSaida_id = #temp.EntradaSaida_id " & vbCrLf &
                    "								   and ni.Serie_id        = #temp.Serie_Id " & vbCrLf &
                    "								   and ni.Nota_Id         = #temp.Nota_id) As ValorFaturado " & vbCrLf &
                    " into #ResumoPorNota" & vbCrLf &
                    " from #temp" & vbCrLf &
                    " " & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    " -- 5 Resultado do resumo por Representante " & vbCrLf &
                    "select #temp.Representante, #temp.EndRepresentante, isnull(rep.Nome,'') AS NomeRepresentante, isnull((ltrim(rep.Cidade) + ' / ' + rep.Estado),'') AS CidadeRepresentante, " & vbCrLf &
                    "		round(sum((#temp.ValorLiquido - isnull(DevolvidoNoPagar.ValorLiquido,0)) * #temp.Percentual / 100),2) AS ValorLiberado, " & vbCrLf &
                    "		isnull(VlrNoPagar.ValorLiquido,0) AS AdiantamentoRepresentante, " & vbCrLf &
                    "		case" & vbCrLf &
                    "			when isnull(VlrNoPagar.ValorLiquido,0) > 0" & vbCrLf &
                    "				then sum(((#temp.ValorLiquido - isnull(DevolvidoNoPagar.ValorLiquido,0)) * #temp.Percentual / 100) - isnull(VlrNoPagar.ValorLiquido,0))" & vbCrLf &
                    "				else sum((#temp.ValorLiquido - isnull(DevolvidoNoPagar.ValorLiquido,0)) * #temp.Percentual / 100)" & vbCrLf &
                    "		end AS SaldoApagar" & vbCrLf &
                    "from #temp" & vbCrLf &
                    "		left join Clientes rep" & vbCrLf &
                    "				on rep.Cliente_id   = #temp.Representante" & vbCrLf &
                    "				and rep.Endereco_id = #temp.EndRepresentante" & vbCrLf &
                    "		left join (select cpg.Empresa, cpg.EndEmpresa, cpg.Cliente, cpg.EndCliente," & vbCrLf &
                    "						sum(cpg.ValorLiquido) as ValorLiquido" & vbCrLf &
                    "					from ContasAPagar cpg" & vbCrLf &
                    "						inner join Pedidos ped" & vbCrLf &
                    "								on ped.Empresa_Id     = cpg.EmpresaPedido" & vbCrLf &
                    "								and ped.EndEmpresa_Id = cpg.EndEmpresaPedido" & vbCrLf &
                    "								and ped.Pedido_Id     = cpg.Pedido" & vbCrLf &
                    "						inner join Comissoes com" & vbCrLf &
                    "								on com.Empresa_Id     = ped.Empresa_Id" & vbCrLf &
                    "								and com.EndEmpresa_Id = ped.EndEmpresa_Id" & vbCrLf &
                    "								and com.Pedido_Id     = ped.Pedido_Id" & vbCrLf &
                    "					 where (cpg.Situacao = 1) " & vbCrLf &
                    "						AND (cpg.Grupado <> 'M') " & vbCrLf &
                    "						AND (cpg.Baixa BETWEEN '2016-11-01' and '2016-11-30')" & vbCrLf &
                    "						and cpg.Provisao = 1" & vbCrLf &
                    "					group by cpg.Empresa, cpg.EndEmpresa, cpg.Cliente, cpg.EndCliente) AS DevolvidoNoPagar" & vbCrLf &
                    "				on DevolvidoNoPagar.Empresa     = #temp.Empresa" & vbCrLf &
                    "				and DevolvidoNoPagar.EndEmpresa = #temp.EndEmpresa" & vbCrLf &
                    "				and DevolvidoNoPagar.Cliente    = #temp.Cliente" & vbCrLf &
                    "				and DevolvidoNoPagar.EndCliente = #temp.EndCliente" & vbCrLf &
                    "		left join (select cpg.Empresa, cpg.EndEmpresa, cpg.Cliente, cpg.EndCliente," & vbCrLf &
                    "						sum(cpg.ValorLiquido) as ValorLiquido" & vbCrLf &
                    "					from ContasAPagar cpg" & vbCrLf &
                    "						inner join Pedidos ped" & vbCrLf &
                    "								on ped.Empresa_Id     = cpg.EmpresaPedido" & vbCrLf &
                    "								and ped.EndEmpresa_Id = cpg.EndEmpresaPedido" & vbCrLf &
                    "								and ped.Pedido_Id     = cpg.Pedido" & vbCrLf &
                    "						inner join Comissoes com" & vbCrLf &
                    "								on com.Empresa_Id     = ped.Empresa_Id" & vbCrLf &
                    "								and com.EndEmpresa_Id = ped.EndEmpresa_Id" & vbCrLf &
                    "								and com.Pedido_Id     = ped.Pedido_Id" & vbCrLf &
                    "					 where (cpg.Situacao = 1) " & vbCrLf &
                    "						AND (cpg.Grupado <> 'M') " & vbCrLf &
                    "						AND (cpg.Baixa BETWEEN '2016-11-01' and '2016-11-30')" & vbCrLf &
                    "						and cpg.Provisao = 1" & vbCrLf &
                    "					group by cpg.Empresa, cpg.EndEmpresa, cpg.Cliente, cpg.EndCliente) AS VlrNoPagar" & vbCrLf &
                    "				on VlrNoPagar.Empresa     = #temp.Empresa" & vbCrLf &
                    "				and VlrNoPagar.EndEmpresa = #temp.EndEmpresa" & vbCrLf &
                    "				and VlrNoPagar.Cliente    = #temp.Representante" & vbCrLf &
                    "				and VlrNoPagar.EndCliente = #temp.EndRepresentante" & vbCrLf &
                    "group by #temp.Representante, #temp.EndRepresentante, rep.Nome, rep.Cidade, rep.Estado, " & vbCrLf &
                    "			VlrNoPagar.ValorLiquido" & vbCrLf &
                    "order by  rep.Nome" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    "" & vbCrLf &
                    " -- 6 Resultado do resumo por Representante/Titulo/Nota" & vbCrLf &
                    " select Representante, EndRepresentante, isnull(rep.Nome,'') AS NomeRepresentante, isnull((ltrim(rep.Cidade) + ' / ' + rep.Estado),'') AS CidadeRepresentante," & vbCrLf &
                    "		Cliente, EndCliente, cli.Nome AS NomeCliente, (ltrim(cli.Cidade) + ' / ' + cli.Estado) AS CidadeCliente, Registro_id AS Titulo, Pedido, Nota_Id AS Nota, ValorFaturado, " & vbCrLf &
                    "		Percentual, sum(ValorFaturado * Percentual / 100) AS Comissao, ValorDoDocumento, ValorLiquido," & vbCrLf &
                    "		sum(ValorLiquido * Percentual / 100) AS ValorLiberado, Historico " & vbCrLf &
                    "from #ResumoPorNota" & vbCrLf &
                    "		inner join Clientes cli" & vbCrLf &
                    "				on cli.Cliente_id   = Cliente" & vbCrLf &
                    "				and cli.Endereco_id = EndCliente" & vbCrLf &
                    "		left join Clientes rep" & vbCrLf &
                    "				on rep.Cliente_id   = Representante" & vbCrLf &
                    "				and rep.Endereco_id = EndRepresentante" & vbCrLf &
                    "group by Representante, EndRepresentante, rep.Nome, rep.Cidade, rep.Estado, " & vbCrLf &
                    "			Cliente, EndCliente, cli.Nome, cli.Cidade, cli.Estado, Registro_id, ValorDoDocumento, ValorLiquido, " & vbCrLf &
                    "			Pedido, Nota_Id, ValorFaturado, Percentual, Historico " & vbCrLf &
                    "order by  rep.Nome"

                    Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Comissoes")

                    '*********************************
                    '*** 1 - Resumo dos Representantes
                    '*********************************

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 8
                    Dim columnIndex As Integer = 1

                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("RESUMO")
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                    worksheet.Cells("A1:H1").Merge = True
                    worksheet.Cells("A2:H2").Merge = True
                    worksheet.Cells("A3:H3").Merge = True
                    worksheet.Cells("A4:H4").Merge = True
                    worksheet.Cells("A5:H5").Merge = True
                    worksheet.Cells("A6:H6").Merge = True
                    worksheet.Cells("A7:H7").Merge = True
                    worksheet.Cells("A1:H7").Style.Font.Bold = True

                    worksheet.Cells("A1").Value = objEmpresa.Nome
                    worksheet.Cells("A2").Value = "           " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                    worksheet.Cells("A3").Value = "           CNPJ - " & objEmpresa.CodigoFormatado
                    worksheet.Cells("A5").Value = "           RESUMO DA COMISSÃO POR REPRESENTANTE"
                    worksheet.Cells("A7").Value = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables("Comissoes").Columns
                        If Not col.ColumnName = "EndRepresentante" Then
                            If col.ColumnName = "Representante" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col.ColumnName = "NomeRepresentante" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Nome"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col.ColumnName = "CidadeRepresentante" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Cidade"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col.ColumnName = "ValorLiberado" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Liberado"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col.ColumnName = "AdiantamentoRepresentante" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Adiantamento"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col.ColumnName = "SaldoApagar" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Saldo a Pagar"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            End If

                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 8

                            columnIndex += 1
                        End If
                    Next

                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Comissoes").Columns.Count - 1)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using

                    rowIndex += 1

                    'criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables("Comissoes").Rows
                        If Not String.IsNullOrWhiteSpace(row("Representante")) Then
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables("Comissoes").Columns
                                If Not col.ColumnName = "EndRepresentante" Then
                                    If IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    ElseIf IsDate(row(col.ColumnName)) Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    End If

                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 8

                                    columnIndex += 1
                                End If
                            Next

                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Comissoes").Columns.Count - 1)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If

                            rowIndex += 1
                        End If
                    Next

                    worksheet.Cells(rowIndex, 3).Value = "TOTAL"
                    worksheet.Cells(rowIndex, 3).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 3).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 3).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 4).Formula = "Sum(" + worksheet.Cells(9, 4).Address + ":" + worksheet.Cells(rowIndex - 1, 4).Address + ")"
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 4).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 4).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 4).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 5).Formula = "Sum(" + worksheet.Cells(9, 5).Address + ":" + worksheet.Cells(rowIndex - 1, 5).Address + ")"
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 5).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 5).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 5).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 6).Formula = "Sum(" + worksheet.Cells(9, 6).Address + ":" + worksheet.Cells(rowIndex - 1, 6).Address + ")"
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 6).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 6).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 6).Style.Font.Size = 8

                    'Auto ajuste Colunas
                    worksheet.Cells.AutoFitColumns(0)

                    'Congelando Painel
                    worksheet.View.FreezePanes(9, 1)

                    '******************************************
                    '*** 2 - 'Criando Pastas por Representante
                    '******************************************
                    For Each row As DataRow In ds.Tables("Comissoes").Rows
                        If Not String.IsNullOrWhiteSpace(row("Representante")) AndAlso Not row("Representante") = "03189063000126" Then
                            worksheet = package.Workbook.Worksheets.Add(row("NomeRepresentante"))
                            worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                            rowIndex = 8
                            columnIndex = 1

                            worksheet.Cells("A1:M1").Merge = True
                            worksheet.Cells("A2:M2").Merge = True
                            worksheet.Cells("A3:M3").Merge = True
                            worksheet.Cells("A4:M4").Merge = True
                            worksheet.Cells("A5:M5").Merge = True
                            worksheet.Cells("A6:M6").Merge = True
                            worksheet.Cells("A7:M7").Merge = True
                            worksheet.Cells("A1:A7").Style.Font.Bold = True

                            worksheet.Cells("A1").Value = objEmpresa.Nome
                            worksheet.Cells("A2").Value = "           " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                            worksheet.Cells("A3").Value = "           CNPJ - " & objEmpresa.CodigoFormatado
                            worksheet.Cells("A5").Value = "           REPRESENTANTE: " & Funcoes.FormatarCpfCnpj(row("Representante")) & " " & row("NomeRepresentante") & "- " & row("CidadeRepresentante")
                            worksheet.Cells(5, 1).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells("A7").Value = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                            'criando linha com o cabeçalho da planilha
                            For Each col1 As DataColumn In ds.Tables("Comissoes1").Columns
                                If Not col1.ColumnName = "Representante" And Not col1.ColumnName = "EndRepresentante" And Not col1.ColumnName = "NomeRepresentante" And Not col1.ColumnName = "CidadeRepresentante" And Not col1.ColumnName = "EndCliente" Then
                                    If col1.ColumnName = "NomeCliente" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "Nome"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    ElseIf col1.ColumnName = "CidadeCliente" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "Cidade"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    ElseIf col1.ColumnName = "Titulo" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = col1.ColumnName
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    ElseIf col1.ColumnName = "Nota" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = col1.ColumnName
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    ElseIf col1.ColumnName = "ValorFaturado" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "Valor NF"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    ElseIf col1.ColumnName = "Percentual" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "%"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    ElseIf col1.ColumnName = "Comissao" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "R$ Comissão"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    ElseIf col1.ColumnName = "ValorDoDocumento" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "R$ Título"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    ElseIf col1.ColumnName = "ValorLiquido" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "R$ Líquido"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    ElseIf col1.ColumnName = "ValorLiberado" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "R$ Liberado"
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = col1.ColumnName
                                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    End If

                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 8

                                    columnIndex += 1
                                End If
                            Next

                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Comissoes1").Columns.Count - 5)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            rowIndex += 1

                            For Each dr As DataRow In ds.Tables("Comissoes1").Rows
                                If dr("Representante") = row("Representante") Then

                                    columnIndex = 1

                                    For Each col2 As DataColumn In ds.Tables("Comissoes1").Columns
                                        If Not col2.ColumnName = "Representante" And Not col2.ColumnName = "EndRepresentante" And Not col2.ColumnName = "NomeRepresentante" And Not col2.ColumnName = "CidadeRepresentante" And Not col2.ColumnName = "EndCliente" Then
                                            If IsNumeric(dr(col2.ColumnName)) AndAlso dr(col2.ColumnName).ToString.Contains(",") Then
                                                worksheet.Cells(rowIndex, columnIndex).Value = dr(col2.ColumnName)
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                            ElseIf IsDate(dr(col2.ColumnName)) Then
                                                worksheet.Cells(rowIndex, columnIndex).Value = dr(col2.ColumnName)
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                            Else
                                                worksheet.Cells(rowIndex, columnIndex).Value = dr(col2.ColumnName)
                                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                            End If

                                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 7

                                            columnIndex += 1
                                        End If
                                    Next

                                    If rowIndex Mod 2 = 0 Then
                                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Comissoes1").Columns.Count - 5)
                                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                        End Using
                                    End If

                                    rowIndex += 1
                                End If
                            Next

                            worksheet.Cells(rowIndex, 3).Value = "TOTAL"
                            worksheet.Cells(rowIndex, 3).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, 3).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, 3).Style.Font.Size = 7

                            worksheet.Cells(rowIndex, 7).Formula = "Sum(" + worksheet.Cells(9, 7).Address + ":" + worksheet.Cells(rowIndex - 1, 7).Address + ")"
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(rowIndex, 7).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, 7).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, 7).Style.Font.Size = 7

                            worksheet.Cells(rowIndex, 9).Formula = "Sum(" + worksheet.Cells(9, 9).Address + ":" + worksheet.Cells(rowIndex - 1, 9).Address + ")"
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(rowIndex, 9).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, 9).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, 9).Style.Font.Size = 7

                            worksheet.Cells(rowIndex, 10).Formula = "Sum(" + worksheet.Cells(9, 10).Address + ":" + worksheet.Cells(rowIndex - 1, 10).Address + ")"
                            worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(rowIndex, 10).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, 10).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, 10).Style.Font.Size = 7

                            worksheet.Cells(rowIndex, 11).Formula = "Sum(" + worksheet.Cells(9, 11).Address + ":" + worksheet.Cells(rowIndex - 1, 11).Address + ")"
                            worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(rowIndex, 11).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, 11).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, 11).Style.Font.Size = 7

                            worksheet.Cells(rowIndex, 12).Formula = "Sum(" + worksheet.Cells(9, 12).Address + ":" + worksheet.Cells(rowIndex - 1, 12).Address + ")"
                            worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(rowIndex, 12).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, 12).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, 12).Style.Font.Size = 7

                            'Auto ajuste Colunas
                            worksheet.Cells.AutoFitColumns(0)

                            'Congelando Painel
                            worksheet.View.FreezePanes(9, 1)
                        End If
                    Next

                    '******************************************************************
                    '*** 3 - 'Resultado do resumo por Representante/Cliente/Titulo/Nota
                    '******************************************************************
                    worksheet = package.Workbook.Worksheets.Add("FINANCEIRO")
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                    rowIndex = 8
                    columnIndex = 1

                    worksheet.Cells("A1:L1").Merge = True
                    worksheet.Cells("A2:L2").Merge = True
                    worksheet.Cells("A3:L3").Merge = True
                    worksheet.Cells("A4:L4").Merge = True
                    worksheet.Cells("A5:L5").Merge = True
                    worksheet.Cells("A6:L6").Merge = True
                    worksheet.Cells("A7:L7").Merge = True
                    worksheet.Cells("A1:A7").Style.Font.Bold = True

                    worksheet.Cells("A1").Value = objEmpresa.Nome
                    worksheet.Cells("A2").Value = "           " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                    worksheet.Cells("A3").Value = "           CNPJ - " & objEmpresa.CodigoFormatado
                    worksheet.Cells("A5").Value = "           RELATÓRIO FINANCEIRO"
                    worksheet.Cells(5, 1).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells("A7").Value = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                    'criando linha com o cabeçalho da planilha
                    For Each col1 As DataColumn In ds.Tables("Comissoes1").Columns
                        If Not col1.ColumnName = "EndRepresentante" And Not col1.ColumnName = "EndCliente" Then
                            If col1.ColumnName = "NomeCliente" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Nome"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col1.ColumnName = "NomeRepresentante" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Nome"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col1.ColumnName = "CidadeCliente" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Cidade"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col1.ColumnName = "CidadeRepresentante" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Cidade"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col1.ColumnName = "Titulo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col1.ColumnName
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col1.ColumnName = "Nota" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col1.ColumnName
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            ElseIf col1.ColumnName = "ValorFaturado" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "Valor NF"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col1.ColumnName = "Percentual" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "%"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col1.ColumnName = "Comissao" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Comissão"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col1.ColumnName = "ValorDoDocumento" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Título"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col1.ColumnName = "ValorLiquido" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Líquido"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            ElseIf col1.ColumnName = "ValorLiberado" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "R$ Liberado"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = col1.ColumnName
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            End If

                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 8

                            columnIndex += 1
                        End If
                    Next

                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Comissoes1").Columns.Count - 2)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using

                    rowIndex += 1

                    For Each dr As DataRow In ds.Tables("Comissoes1").Rows
                        columnIndex = 1

                        For Each col2 As DataColumn In ds.Tables("Comissoes1").Columns
                            If Not col2.ColumnName = "EndRepresentante" And Not col2.ColumnName = "EndCliente" Then
                                If IsNumeric(dr(col2.ColumnName)) AndAlso dr(col2.ColumnName).ToString.Contains(",") Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = dr(col2.ColumnName)
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                ElseIf IsDate(dr(col2.ColumnName)) Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = dr(col2.ColumnName)
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = dr(col2.ColumnName)
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                End If

                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 8

                                columnIndex += 1
                            End If
                        Next

                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Comissoes1").Columns.Count - 2)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If

                        rowIndex += 1
                    Next

                    worksheet.Cells(rowIndex, 6).Value = "TOTAL"
                    worksheet.Cells(rowIndex, 6).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 6).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 6).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 10).Formula = "Sum(" + worksheet.Cells(9, 10).Address + ":" + worksheet.Cells(rowIndex - 1, 10).Address + ")"
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 10).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 10).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 10).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 12).Formula = "Sum(" + worksheet.Cells(9, 12).Address + ":" + worksheet.Cells(rowIndex - 1, 12).Address + ")"
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 12).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 12).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 12).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 13).Formula = "Sum(" + worksheet.Cells(9, 13).Address + ":" + worksheet.Cells(rowIndex - 1, 13).Address + ")"
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 13).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 13).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 13).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 14).Formula = "Sum(" + worksheet.Cells(9, 14).Address + ":" + worksheet.Cells(rowIndex - 1, 14).Address + ")"
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 14).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 14).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 14).Style.Font.Size = 8

                    worksheet.Cells(rowIndex, 15).Formula = "Sum(" + worksheet.Cells(9, 15).Address + ":" + worksheet.Cells(rowIndex - 1, 15).Address + ")"
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, 15).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, 15).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, 15).Style.Font.Size = 8

                    'Auto ajuste Colunas
                    worksheet.Cells.AutoFitColumns(0)

                    'Congelando Painel
                    worksheet.View.FreezePanes(9, 1)

                    package.Save()
                End Using
            End Using

            'download do arquivo pelo browser
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try


            If Not ValidaCampos() Then
                Exit Sub
            End If

            'Dim valorFat As Decimal
            'Dim comissao As Decimal
            Dim pedido As Integer
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue().Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")
            Dim ds As DataSet = getDataSetExcel()

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("COMISSÕES")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A1:AC" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    'criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            'If pedido <> row("Pedido") Then
                            If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Contains(",") Then
                                worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                            ElseIf IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            ElseIf IsDate(row(col.ColumnName)) Then
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If

                            If col.ColumnName = "ValorLiberado" Then
                                If row("Operacao") = 36 Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = row("Comissao")
                                End If
                            End If
                            'Else
                            '    If col.ColumnName = "Titulo" OrElse col.ColumnName = "DataTitulo" OrElse col.ColumnName = "ValorPago" OrElse col.ColumnName = "ValorLiberado" Then
                            '        If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Contains(",") Then
                            '            worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                            '            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                            '        ElseIf IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                            '            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            '            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            '        ElseIf IsDate(row(col.ColumnName)) Then
                            '            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            '            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                            '        Else
                            '            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            '        End If
                            '    End If
                            'End If

                            columnIndex += 1
                        Next

                        pedido = row("Pedido")
                        'aplicando formatação decimal nos campos de valores
                        'worksheet.Cells(String.Format("S{0}:S{0}", rowIndex)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(2, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using

            'download do arquivo pelo browser
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

#End Region

    Protected Sub lnkResumoExcel_Click(sender As Object, e As EventArgs) Handles lnkResumoExcel.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Empresa não foi selecionada", eTitulo.Info)
            Else
                resumoExcel()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeComissoesPedido")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class