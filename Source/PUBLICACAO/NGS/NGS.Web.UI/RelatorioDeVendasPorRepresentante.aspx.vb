Imports System.Drawing
Imports System.IO
Imports NGS.Lib.Negocio
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class RelatorioDeVendasPorRepresentante
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeVendasPorRepresentante", "ACESSAR") Then
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = ""
        If rbPorRepresentante.Checked Then
            sql &= " Select * from (" & vbCrLf &
                            " 		 SELECT nf.Empresa_Id, nf.EndEmpresa_Id, 'Mercado Nacional' as Mercado, 1 as Ordem, " & vbCrLf &
                            " 		        Clientes.Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado, SUM(CASE WHEN MONTH(Movimento) = 1 THEN nfXItens.Valor ELSE 0 END)  AS Janeiro, " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 2 THEN nfXItens.Valor ELSE 0 END)  AS Fevereiro,                         " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 3 THEN nfXItens.Valor ELSE 0 END)  AS Marco,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 4 THEN nfXItens.Valor ELSE 0 END)  AS Abril,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 5 THEN nfXItens.Valor ELSE 0 END)  AS Maio,                              " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 6 THEN nfXItens.Valor ELSE 0 END)  AS Junho,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 7 THEN nfXItens.Valor ELSE 0 END)  AS julho,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 8 THEN nfXItens.Valor ELSE 0 END)  AS Agosto,                            " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 9 THEN nfXItens.Valor ELSE 0 END)  AS Setembro,                          " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 10 THEN nfXItens.Valor ELSE 0 END) AS Outubro,                           " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 11 THEN nfXItens.Valor ELSE 0 END) AS Novembro,                          " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 12 THEN nfXItens.Valor ELSE 0 END) AS Dezembro,                          " & vbCrLf &
                            "                 SUM(nfXItens.Valor) AS Acumulado,                                                                         " & vbCrLf &
                            "                 SUM(nfXItens.QuantidadeFiscal) AS QuantidadeFiscal                                                        " & vbCrLf &
                            "            FROM NotasFiscais nf                                                                                           " & vbCrLf &
                            " 		  INNER JOIN Comissoes                                                                                              " & vbCrLf &
                            " 		     ON nf.Empresa_Id = Comissoes.Empresa_Id                                                                        " & vbCrLf &
                            " 		    AND nf.EndEmpresa_Id = Comissoes.EndEmpresa_Id                                                                  " & vbCrLf &
                            " 		    AND nf.Pedido = Comissoes.Pedido_Id                                                                             " & vbCrLf &
                            " 		  INNER JOIN NotasFiscaisXItens nfXItens                                                                            " & vbCrLf &
                            " 		     ON nf.Empresa_Id = nfXItens.Empresa_Id                                                                         " & vbCrLf &
                            " 		    AND nf.EndEmpresa_Id = nfXItens.EndEmpresa_Id                                                                   " & vbCrLf &
                            " 			AND nf.Cliente_Id = nfXItens.Cliente_Id                                                                         " & vbCrLf &
                            " 		    AND nf.EndCliente_Id = nfXItens.EndCliente_Id                                                                   " & vbCrLf &
                            " 			AND nf.EntradaSaida_Id = nfXItens.EntradaSaida_Id                                                               " & vbCrLf &
                            " 			AND nf.Serie_Id = nfXItens.Serie_Id                                                                             " & vbCrLf &
                            " 			AND nf.Nota_Id  = nfXItens.Nota_Id                                                                              " & vbCrLf &
                            " 		  INNER JOIN Clientes                                                                                               " & vbCrLf &
                            " 			 ON Comissoes.Representante_Id    = Clientes.Cliente_Id                                                         " & vbCrLf &
                            " 		    AND Comissoes.EndRepresentante_Id = Clientes.Endereco_Id                                                        " & vbCrLf &
                            " 		  INNER JOIN SubOperacoes so                                                                                        " & vbCrLf &
                            " 			 ON nf.Operacao    = so.Operacao_Id                                                                             " & vbCrLf &
                            " 			AND nf.SubOperacao = so.SubOperacoes_Id                                                                         " & vbCrLf &
                            "           WHERE YEAR(nf.Movimento) = " & DdlAno.SelectedValue & vbCrLf &
                            "             And MONTH(nf.movimento) <= " & ddlMesBase.SelectedValue & vbCrLf &
                            "             And so.Classe in ('" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.VENDASAORDEM.ToString & "')" & vbCrLf &
                            "           GROUP BY nf.Empresa_Id, nf.EndEmpresa_Id, Clientes.Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf &
                            "           Union                                                                                                           " & vbCrLf &
                            "          SELECT nf.Empresa_Id,                                                                                            " & vbCrLf &
                            "                 nf.EndEmpresa_Id,                                                                                         " & vbCrLf &
                            "                 '" & eClassesOperacoes.EXPORTACOES.ToString & "' As Mercado, 2 as Ordem,                                  " & vbCrLf &
                            "                 Clientes.Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado,                                     " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 1  THEN nfXItens.Valor ELSE 0 END)  AS Janeiro,                           " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 2  THEN nfXItens.Valor ELSE 0 END)  AS Fevereiro,                         " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 3  THEN nfXItens.Valor ELSE 0 END)  AS Marco,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 4  THEN nfXItens.Valor ELSE 0 END)  AS Abril,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 5  THEN nfXItens.Valor ELSE 0 END)  AS Maio,                              " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 6  THEN nfXItens.Valor ELSE 0 END)  AS Junho,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 7  THEN nfXItens.Valor ELSE 0 END)  AS julho,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 8  THEN nfXItens.Valor ELSE 0 END)  AS Agosto,                            " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 9  THEN nfXItens.Valor ELSE 0 END)  AS Setembro,                          " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 10 THEN nfXItens.Valor ELSE 0 END) AS Outubro,                           " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 11 THEN nfXItens.Valor ELSE 0 END) AS Novembro,                          " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 12 THEN nfXItens.Valor ELSE 0 END) AS desembro,                          " & vbCrLf &
                            "                 SUM(nfXItens.Valor) AS Acumulado,                                                                         " & vbCrLf &
                            "                 SUM(nfXItens.QuantidadeFiscal) AS QuantidadeFiscal                                                        " & vbCrLf &
                            "            FROM NotasFiscais nf                                                                                           " & vbCrLf &
                            " 		  INNER JOIN Comissoes                                                                                              " & vbCrLf &
                            " 			 ON nf.Empresa_Id = Comissoes.Empresa_Id                                                                        " & vbCrLf &
                            " 			AND nf.EndEmpresa_Id = Comissoes.EndEmpresa_Id                                                                  " & vbCrLf &
                            " 			AND nf.Pedido = Comissoes.Pedido_Id                                                                             " & vbCrLf &
                            " 		  INNER JOIN NotasFiscaisXItens nfXItens                                                                            " & vbCrLf &
                            " 			 ON nf.Empresa_Id = nfXItens.Empresa_Id                                                                         " & vbCrLf &
                            " 			AND nf.EndEmpresa_Id = nfXItens.EndEmpresa_Id                                                                   " & vbCrLf &
                            " 			AND nf.Cliente_Id = nfXItens.Cliente_Id                                                                         " & vbCrLf &
                            " 			AND nf.EndCliente_Id = nfXItens.EndCliente_Id                                                                   " & vbCrLf &
                            " 			AND nf.EntradaSaida_Id = nfXItens.EntradaSaida_Id                                                               " & vbCrLf &
                            " 			AND nf.Serie_Id = nfXItens.Serie_Id                                                                             " & vbCrLf &
                            " 			AND nf.Nota_Id = nfXItens.Nota_Id                                                                               " & vbCrLf &
                            " 		  INNER JOIN Clientes                                                                                               " & vbCrLf &
                            " 			 ON Comissoes.Representante_Id = Clientes.Cliente_Id                                                            " & vbCrLf &
                            " 			AND Comissoes.EndRepresentante_Id = Clientes.Endereco_Id                                                        " & vbCrLf &
                            " 		  INNER JOIN SubOperacoes so                                                                                        " & vbCrLf &
                            " 		     ON nf.Operacao = so.Operacao_Id                                                                                " & vbCrLf &
                            " 			AND nf.SubOperacao = so.suboperacoes_Id                                                                         " & vbCrLf &
                            "           WHERE YEAR(nf.Movimento) = " & DdlAno.SelectedValue & " And so.Classe in ('" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf &
                            "             And MONTH(nf.movimento) <= " & ddlMesBase.SelectedValue & vbCrLf &
                            "           GROUP BY nf.Empresa_Id, nf.EndEmpresa_Id, Clientes.Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado  " & vbCrLf &
                            "           Union                                                                                                           " & vbCrLf &
                            "          SELECT nf.Empresa_Id,                                                                                            " & vbCrLf &
                            "                 nf.EndEmpresa_Id,                                                                                         " & vbCrLf &
                            "                         'CONSOLIDADO' As Mercado, 3 as Ordem,                                                             " & vbCrLf &
                            "                 Clientes.Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado,                                     " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 1 THEN nfXItens.Valor ELSE 0 END)  AS Janeiro,                           " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 2 THEN nfXItens.Valor ELSE 0 END)  AS Fevereiro,                         " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 3 THEN nfXItens.Valor ELSE 0 END)  AS Marco,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 4 THEN nfXItens.Valor ELSE 0 END)  AS Abril,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 5 THEN nfXItens.Valor ELSE 0 END)  AS Maio,                              " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 6 THEN nfXItens.Valor ELSE 0 END)  AS Junho,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 7 THEN nfXItens.Valor ELSE 0 END)  AS julho,                             " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 8 THEN nfXItens.Valor ELSE 0 END)  AS Agosto,                            " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 9 THEN nfXItens.Valor ELSE 0 END)  AS Setembro,                          " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 10 THEN nfXItens.Valor ELSE 0 END) AS Outubro,                           " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 11 THEN nfXItens.Valor ELSE 0 END) AS Novembro,                          " & vbCrLf &
                            "                 SUM(CASE WHEN MONTH(Movimento) = 12 THEN nfXItens.Valor ELSE 0 END) AS desembro,                          " & vbCrLf &
                            "                 SUM(nfXItens.Valor) AS Acumulado,                                                                         " & vbCrLf &
                            "                 SUM(nfXItens.QuantidadeFiscal) AS QuantidadeFiscal                                                        " & vbCrLf &
                            "            FROM NotasFiscais nf                                                                                           " & vbCrLf &
                            " 		  INNER JOIN Comissoes                                                                                              " & vbCrLf &
                            " 		     ON nf.Empresa_Id = Comissoes.Empresa_Id                                                                        " & vbCrLf &
                            " 			AND nf.EndEmpresa_Id = Comissoes.EndEmpresa_Id                                                                  " & vbCrLf &
                            " 			AND nf.Pedido = Comissoes.Pedido_Id                                                                             " & vbCrLf &
                            " 		  INNER JOIN NotasFiscaisXItens nfXItens                                                                            " & vbCrLf &
                            " 		     ON nf.Empresa_Id = nfXItens.Empresa_Id                                                                         " & vbCrLf &
                            " 		    AND nf.EndEmpresa_Id = nfXItens.EndEmpresa_Id                                                                   " & vbCrLf &
                            " 			AND nf.Cliente_Id = nfXItens.Cliente_Id                                                                         " & vbCrLf &
                            " 			AND nf.EndCliente_Id = nfXItens.EndCliente_Id                                                                   " & vbCrLf &
                            " 			AND nf.EntradaSaida_Id = nfXItens.EntradaSaida_Id                                                               " & vbCrLf &
                            " 			AND nf.Serie_Id = nfXItens.Serie_Id                                                                             " & vbCrLf &
                            " 			AND nf.Nota_Id = nfXItens.Nota_Id                                                                               " & vbCrLf &
                            " 		  INNER JOIN Clientes                                                                                               " & vbCrLf &
                            " 		     ON Comissoes.Representante_Id = Clientes.Cliente_Id                                                            " & vbCrLf &
                            " 			AND Comissoes.EndRepresentante_Id = Clientes.Endereco_Id                                                        " & vbCrLf &
                            " 		  INNER JOIN SubOperacoes so                                                                                        " & vbCrLf &
                            " 		     ON nf.Operacao = so.Operacao_Id                                                                                " & vbCrLf &
                            " 			AND nf.SubOperacao = so.suboperacoes_Id                                                                         " & vbCrLf &
                            "           WHERE YEAR(nf.Movimento) = " & DdlAno.SelectedValue & "                                                         " & vbCrLf &
                            "             And MONTH(nf.movimento) <= " & ddlMesBase.SelectedValue & vbCrLf &
                            "           GROUP BY nf.Empresa_Id, nf.EndEmpresa_Id, Clientes.Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado  " & vbCrLf &
                            "          ) as Consulta                                                                                                    " & vbCrLf

            If chkEmpresaCons.Checked Then
                sql &= " where left(Consulta.Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= " where Consulta.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If txtCodigoCliente.Value.ToString.Length > 0 Then
                sql &= "   And Consulta.Cliente_Id = '" & txtCodigoCliente.Value.ToString.Split("-")(0) & "'" & vbCrLf
            End If

            sql &= " ORDER BY Ordem, Nome                                                                                                      " & vbCrLf
        Else
            sql &= " Select Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                    "       Notas.Cliente_Id, Notas.EndCliente_Id, Mercado, Ordem, Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf &
                    "       SUM(Acumulado) as Acumulado, " & vbCrLf &
                    "       Sum(Vencidos2015) as Vencidos2015, " & vbCrLf &
                    "       Sum(Vencidos2016) as Vencidos2016, " & vbCrLf &
                    "       Sum(Vencidos2015) + Sum(Vencidos2016) as Vencidos " & vbCrLf &
                    "  From ( " & vbCrLf &
                    "        SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, 'Mercado Nacional' AS Mercado, 1 AS Ordem, " & vbCrLf &
                    "               0 AS Acumulado, " & vbCrLf &
                    "               SUM(CASE WHEN year(NotasFiscais.Movimento) < " & DdlAno.SelectedValue & " THEN ContasAReceber.ValorLiquido ELSE 0 END) AS Vencidos2015, " & vbCrLf &
                    "               SUM(CASE WHEN year(NotasFiscais.Movimento) = " & DdlAno.SelectedValue & " THEN ContasAReceber.ValorLiquido ELSE 0 END) AS Vencidos2016  " & vbCrLf &
                    "          FROM NotasFiscais" & vbCrLf &
                    "         INNER JOIN NotaFiscalXTitulo" & vbCrLf &
                    "            ON NotasFiscais.Empresa_Id      = NotaFiscalXTitulo.Empresa_Id" & vbCrLf &
                    "           AND NotasFiscais.EndEmpresa_Id   = NotaFiscalXTitulo.EndEmpresa_Id" & vbCrLf &
                    "           AND NotasFiscais.Cliente_Id      = NotaFiscalXTitulo.Cliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EndCliente_Id   = NotaFiscalXTitulo.EndCliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EntradaSaida_Id = NotaFiscalXTitulo.EntradaSaida_Id" & vbCrLf &
                    "           AND NotasFiscais.Serie_Id        = NotaFiscalXTitulo.Serie_Id" & vbCrLf &
                    "           AND NotasFiscais.Nota_Id         = NotaFiscalXTitulo.Nota_Id" & vbCrLf &
                    "         INNER JOIN ContasAReceber" & vbCrLf &
                    "            ON NotaFiscalXTitulo.Titulo_Id = ContasAReceber.Registro_Id" & vbCrLf &
                    "         INNER JOIN SubOperacoes" & vbCrLf &
                    "            ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                    "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                    "         WHERE (ContasAReceber.Provisao = 2)" & vbCrLf &
                    "           And (ContasAReceber.Situacao = 1)" & vbCrLf &
                    "           AND SubOperacoes.Classe IN ('" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.VENDASAORDEM.ToString & "') " & vbCrLf &
                    "         GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id" & vbCrLf &
                    "         Union all" & vbCrLf &
                    "        SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, 'Mercado Nacional' AS Mercado, 1 AS Ordem," & vbCrLf &
                    "               SUM(NotasFiscaisXItens.Valor) AS Acumulado," & vbCrLf &
                    "               0  as Vencidos2015, 0  as Vencidos2016 " & vbCrLf &
                    "          FROM NotasFiscais" & vbCrLf &
                    "         INNER JOIN NotasFiscaisXItens" & vbCrLf &
                    "            ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                    "           AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                    "           AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
                    "           AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                    "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
                    "           AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                    "           AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                    "         INNER JOIN SubOperacoes" & vbCrLf &
                    "            ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf &
                    "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                    "         WHERE (Year(NotasFiscais.Movimento) = " & DdlAno.SelectedValue & vbCrLf &
                    "           And Month(NotasFiscais.movimento) <= " & ddlMesBase.SelectedValue & ")" & vbCrLf &
                    "           And  SubOperacoes.Classe IN ('" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.VENDASAORDEM.ToString & "')" & vbCrLf &
                    "         GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id" & vbCrLf &
                    "         Union all" & vbCrLf &
                    "        SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, '" & eClassesOperacoes.EXPORTACOES.ToString & "' AS Mercado, 2 AS Ordem, " & vbCrLf &
                    "               0 AS Acumulado, " & vbCrLf &
                    "               SUM(CASE WHEN year(NotasFiscais.Movimento) < " & DdlAno.SelectedValue & " THEN ContasAReceber.ValorLiquido ELSE 0 END) AS Vencidos2015, " & vbCrLf &
                    "               SUM(CASE WHEN year(NotasFiscais.Movimento) = " & DdlAno.SelectedValue & " THEN ContasAReceber.ValorLiquido ELSE 0 END) AS Vencidos2016" & vbCrLf &
                    "          FROM NotasFiscais" & vbCrLf &
                    "         INNER JOIN NotaFiscalXTitulo" & vbCrLf &
                    "            ON NotasFiscais.Empresa_Id      = NotaFiscalXTitulo.Empresa_Id" & vbCrLf &
                    "           AND NotasFiscais.EndEmpresa_Id   = NotaFiscalXTitulo.EndEmpresa_Id" & vbCrLf &
                    "           AND NotasFiscais.Cliente_Id      = NotaFiscalXTitulo.Cliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EndCliente_Id   = NotaFiscalXTitulo.EndCliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EntradaSaida_Id = NotaFiscalXTitulo.EntradaSaida_Id" & vbCrLf &
                    "           AND NotasFiscais.Serie_Id        = NotaFiscalXTitulo.Serie_Id" & vbCrLf &
                    "           AND NotasFiscais.Nota_Id         = NotaFiscalXTitulo.Nota_Id" & vbCrLf &
                    "         INNER JOIN ContasAReceber" & vbCrLf &
                    "            ON NotaFiscalXTitulo.Titulo_Id = ContasAReceber.Registro_Id" & vbCrLf &
                    "         INNER JOIN SubOperacoes" & vbCrLf &
                    "            ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                    "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                    "         WHERE (ContasAReceber.Provisao = 2)" & vbCrLf &
                    "           And (ContasAReceber.Situacao = 1)" & vbCrLf &
                    "           AND SubOperacoes.Classe IN ('" & eClassesOperacoes.EXPORTACOES.ToString & "') " & vbCrLf &
                    "         GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id" & vbCrLf &
                    "         Union all" & vbCrLf &
                    "        SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, '" & eClassesOperacoes.EXPORTACOES.ToString & "' AS Mercado," & vbCrLf &
                    "               2 AS Ordem,  " & vbCrLf &
                    "               SUM(NotasFiscaisXItens.Valor) AS Acumulado," & vbCrLf &
                    "               0 as Vencidos2015," & vbCrLf &
                    "               0 as Vencidos2016" & vbCrLf &
                    "          FROM NotasFiscais" & vbCrLf &
                    "         INNER JOIN NotasFiscaisXItens" & vbCrLf &
                    "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                    "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                    "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                    "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                    "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                    "         INNER JOIN SubOperacoes" & vbCrLf &
                    "            ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                    "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                    "         WHERE Year(NotasFiscais.Movimento)   = " & DdlAno.SelectedValue &
                    "           AND Month(NotasFiscais.movimento) <= " & ddlMesBase.SelectedValue & vbCrLf &
                    "           And SubOperacoes.Classe IN ('" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf &
                    "         GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id " & vbCrLf &
                    "---------Consolidado-------------------------------------------- " & vbCrLf &
                    "         Union all" & vbCrLf &
                    "        SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, 'Consolidado' AS Mercado, 3 AS Ordem, " & vbCrLf &
                    "               0 AS Acumulado, " & vbCrLf &
                    "               SUM(CASE WHEN year(NotasFiscais.Movimento) < " & DdlAno.SelectedValue & " THEN ContasAReceber.ValorLiquido ELSE 0 END) AS Vencidos2015, " & vbCrLf &
                    "               SUM(CASE WHEN year(NotasFiscais.Movimento) = " & DdlAno.SelectedValue & " THEN ContasAReceber.ValorLiquido ELSE 0 END) AS Vencidos2016 " & vbCrLf &
                    "          FROM NotasFiscais " & vbCrLf &
                    "         INNER JOIN NotaFiscalXTitulo" & vbCrLf &
                    "            ON NotasFiscais.Empresa_Id      = NotaFiscalXTitulo.Empresa_Id" & vbCrLf &
                    "           AND NotasFiscais.EndEmpresa_Id   = NotaFiscalXTitulo.EndEmpresa_Id" & vbCrLf &
                    "           AND NotasFiscais.Cliente_Id      = NotaFiscalXTitulo.Cliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EndCliente_Id   = NotaFiscalXTitulo.EndCliente_Id" & vbCrLf &
                    "           AND NotasFiscais.EntradaSaida_Id = NotaFiscalXTitulo.EntradaSaida_Id" & vbCrLf &
                    "           AND NotasFiscais.Serie_Id        = NotaFiscalXTitulo.Serie_Id" & vbCrLf &
                    "           AND NotasFiscais.Nota_Id         = NotaFiscalXTitulo.Nota_Id" & vbCrLf &
                    "         INNER JOIN ContasAReceber" & vbCrLf &
                    "            ON NotaFiscalXTitulo.Titulo_Id = ContasAReceber.Registro_Id" & vbCrLf &
                    "         INNER JOIN SubOperacoes" & vbCrLf &
                    "            ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id" & vbCrLf &
                    "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                    "         WHERE ContasAReceber.Provisao = 2" & vbCrLf &
                    "           AND ContasAReceber.Situacao = 1" & vbCrLf &
                    "           AND SubOperacoes.Classe IN ('" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.VENDASAORDEM.ToString & "', '" & eClassesOperacoes.EXPORTACOES.ToString & "')  " & vbCrLf &
                    "         GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id " & vbCrLf &
                    "         Union all" & vbCrLf &
                    "        SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, 'Consolidado' AS Mercado, " & vbCrLf &
                    "               3 AS Ordem, " & vbCrLf &
                    "               SUM(NotasFiscaisXItens.Valor) AS Acumulado," & vbCrLf &
                    "               0  as Vencidos2015," & vbCrLf &
                    "               0  as Vencidos2016  " & vbCrLf &
                    "          FROM NotasFiscais" & vbCrLf &
                    "         INNER JOIN NotasFiscaisXItens  " & vbCrLf &
                    "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf &
                    "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                    "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
                    "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                    "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                    "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                    "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                    "         INNER JOIN SubOperacoes " & vbCrLf &
                    "            ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                    "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                    "         WHERE Year(NotasFiscais.Movimento)   = " & DdlAno.SelectedValue &
                    "           And Month(NotasFiscais.movimento) <= " & ddlMesBase.SelectedValue & vbCrLf &
                    "           And  SubOperacoes.Classe IN ('" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.VENDASAORDEM.ToString & "', '" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf &
                    "         GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id" & vbCrLf &
                    "       ) as Notas " & vbCrLf &
                    "  INNER JOIN Clientes  " & vbCrLf &
                    "     ON Notas.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf &
                    "    And Notas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf

            If chkEmpresaCons.Checked Then
                sql &= " where left(Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= " where Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If txtCodigoCliente.Value.ToString.Length > 0 Then
                sql &= "   And Notas.Cliente_Id = '" & txtCodigoCliente.Value.ToString.Split("-")(0) & "'" & vbCrLf
            End If

            sql &= " Group By  Empresa_Id, EndEmpresa_Id, Notas.Cliente_Id, Notas.EndCliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Mercado, Ordem" & vbCrLf &
                   " Order By  Empresa_Id, EndEmpresa_Id, Ordem, Clientes.Nome  " & vbCrLf

        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, IIf(rbPorRepresentante.Checked, "VendasPorRepresentante", "VendasPorCliente"))
        Return ds
    End Function

    Private Function getDataSetPorPedido()
        Dim sql As String = "Select p.empresa_id as empresa, emp.nome + ' - ' + emp.cidade + '/'+ emp.estado as nomeempresa, p.pedido_id as Pedido," & vbCrLf &
                            "       p.cliente, c.nome as nomecliente, convert(varchar,p.datapedido,103) as datapedido, p.operacao, p.suboperacao, so.descricao," & vbCrLf &
                            "	    sum(case when pxixl.tipodelancamento = 'E' then pxixl.totaloficial * -1 else pxixl.totaloficial end) as Total," & vbCrLf &
                            "	    case when isnull(rep.nome,'') = ''" & vbCrLf &
                            "		       then 'NÃO INFORMADO'" & vbCrLf &
                            "			   else rep.cliente_id + '-' + convert(varchar,rep.endereco_id) + ' - ' + rep.nome" & vbCrLf &
                            "	    end as representante, co.percentual" & vbCrLf &
                            "  from pedidos p" & vbCrLf &
                            " inner join pedidoxitem pxi" & vbCrLf &
                            "    on pxi.empresa_id    = p.empresa_id" & vbCrLf &
                            "   and pxi.endempresa_id = p.endempresa_id" & vbCrLf &
                            "   and pxi.pedido_id     = p.pedido_id" & vbCrLf &
                            " inner join pedidoxitemxlancamento pxixl" & vbCrLf &
                            "    on pxixl.empresa_id = pxi.Empresa_Id" & vbCrLf &
                            "   and pxixl.EndEmpresa_Id = pxi.EndEmpresa_Id" & vbCrLf &
                            "   and pxixl.Pedido_Id     = pxi.Pedido_Id" & vbCrLf &
                            "   and pxixl.produto_id    = pxi.produto_id" & vbCrLf &
                            " inner join suboperacoes so" & vbCrLf &
                            "	  on so.operacao_id    = p.operacao" & vbCrLf &
                            "   and so.suboperacoes_id = p.suboperacao" & vbCrLf &
                            "  left join comissoes co" & vbCrLf &
                            "	  on co.empresa_id   = p.empresa_id" & vbCrLf &
                            "   and co.endempresa_id = p.endempresa_id" & vbCrLf &
                            "   and co.pedido_id     = p.pedido_id" & vbCrLf &
                            "  left join clientes emp" & vbCrLf &
                            "    on emp.cliente_id  = p.empresa_id" & vbCrLf &
                            "   and emp.endereco_id = p.endempresa_id" & vbCrLf &
                            "  left join clientes c" & vbCrLf &
                            "    on c.cliente_id  = p.cliente" & vbCrLf &
                            "   and c.endereco_id = p.endcliente" & vbCrLf &
                            "  left join clientes rep" & vbCrLf &
                            "    on rep.cliente_id  = co.Representante_Id" & vbCrLf &
                            "   and rep.endereco_id = co.EndRepresentante_Id" & vbCrLf &
                            " where so.entradasaida = 'S'" & vbCrLf &
                            "   and so.pedido       = 1" & vbCrLf &
                            "   and so.classe       in('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf &
                            "   and p.datapedido    between '" & txtData1.Text.ToSqlDate() & "' and '" & txtData2.Text.ToSqlDate() & "'" & vbCrLf &
                            "   and p.situacao      = 1" & vbCrLf

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            sql &= "   And rep.cliente_id = '" & txtCodigoCliente.Value.ToString.Split("-")(0) & "'" & vbCrLf
        End If

        If chkRepresentanteNaoInformado.Checked Then
            sql &= " and isnull(c.nome,'') = ''    " & vbCrLf
        End If

        sql &= " Group by p.empresa_id, emp.nome + ' - ' + emp.cidade + '/'+ emp.estado, p.pedido_id, p.cliente, c.nome," & vbCrLf &
               "          convert(varchar,p.datapedido,103), p.operacao, p.suboperacao, so.descricao," & vbCrLf &
               "          case when isnull(rep.nome,'') = ''" & vbCrLf &
               " 		    then 'NÃO INFORMADO'" & vbCrLf &
               " 			else rep.cliente_id + '-' + convert(varchar,rep.endereco_id) + ' - ' + rep.nome" & vbCrLf &
               " 	      end, co.percentual" & vbCrLf &
               " Order by p.empresa_id, p.pedido_id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "VendasPorPedido")
        Return ds
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try

            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a empresa.")
            ElseIf Not rbPorPedido.Checked AndAlso (String.IsNullOrWhiteSpace(DdlAno.SelectedValue) OrElse Not IsNumeric(DdlAno.SelectedValue)) Then
                MsgBox(Me.Page, "Informe um exercício válido.")
            ElseIf Not rbPorPedido.Checked AndAlso CInt(DdlAno.SelectedValue) < 2016 Then
                MsgBox(Me.Page, "Exercício de ve ser maior ou igual a 2016.")
            ElseIf rbPorRepresentante.Checked OrElse rbPorCliente.Checked Then

                Dim param As Dictionary(Of String, Object) = getParametrosConsulta(False)

                Dim ds As DataSet = getDataSet()
                Funcoes.BindReport(Me.Page, ds, IIf(rbPorRepresentante.Checked, "Cr_VendasPorRepresentante", "Cr_VendasPorCliente"), eExportType.PDF, param)
            Else
                If String.IsNullOrEmpty(txtData1.Text) OrElse String.IsNullOrEmpty(txtData2.Text) Then
                    MsgBox(Me.Page, "Período é obrigatório.")
                Else
                    Dim param As Dictionary(Of String, Object) = getParametrosConsulta(True)

                    Dim ds As DataSet = getDataSetPorPedido()
                    Funcoes.BindReport(Me.Page, ds, "Cr_VendasPorPedido", eExportType.PDF, param)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorioDados() 'Excel Dados
    End Sub

    Private Sub EmitirRelatorioDados()
        Try
            Dim Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
            Dim data As String = CStr(ddlMesBase.SelectedValue & "/" & DdlAno.SelectedValue)
            Dim ds As DataSet = getDataSet()

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)


                    'criando planilha 
                    Dim worksheet As ExcelWorksheet

                    'criando título da planilha 
                    If (chkRepresentanteNaoInformado.Checked) Then
                        worksheet = package.Workbook.Worksheets.Add("Relatório de vendas por representante.")
                    Else
                        worksheet = package.Workbook.Worksheets.Add("Relatório de vendas.")
                    End If

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
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
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELATÓRIO DE VENDAS")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & data)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        If col.ColumnName <> "Row" Then
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        End If
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células valores
                        worksheet.Cells(String.Format("I{0}:K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParametrosConsulta(ByVal porpedido As Boolean) As Dictionary(Of String, Object)
        Dim obj As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
        Dim param As New Dictionary(Of String, Object)
        param.Add("parametrosconsulta", "Unidade: " & ddlUnidade.SelectedItem.Text & vbCrLf)
        param("parametrosconsulta") &= "Empresa" & IIf(chkEmpresaCons.Checked, " Consolidada: " & obj.Nome, ": ") & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " / " & obj.CodigoEstado & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Dim objCliente = New Cliente(txtCodigoCliente.Value.ToString.Split("-")(0), txtCodigoCliente.Value.ToString.Split("-")(1))
            param("parametrosconsulta") &= "Representante: " & Funcoes.FormatarCpfCnpj(objCliente.Codigo) & " - " & objCliente.Nome
        End If

        If porpedido Then
            param("parametrosconsulta") &= " Período de: " & txtData1.Text & " á " & txtData2.Text & vbCrLf
            param("parametrosconsulta") &= IIf(chkRepresentanteNaoInformado.Checked, "Emissão de pedido sem representante", "")
        Else
            param("parametrosconsulta") &= " Mês Base: " & ddlMesBase.SelectedItem.Text & " de " & DdlAno.SelectedValue
        End If

        Return param
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteRELREP" & HID.Value) Is Nothing Then
            txtCodigoCliente.Value = CType(Session("objClienteRELREP" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objClienteRELREP" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(CType(Session("objClienteRELREP" & HID.Value), [Lib].Negocio.Cliente).Codigo) & " - " & CType(Session("objClienteRELREP" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteRELREP" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
        DdlAno.SelectedValue = Year(Now)
        txtData1.Text = New DateTime(Now.Year, Now.Month, 1)
        txtData2.Text = New DateTime(Now.Year, Now.Month, Now.Day)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteRELREP" & HID.Value, "txtNome")
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeVendasPorRepresentante")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub
End Class