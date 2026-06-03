Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ComposicaoDoEstoqueRN
    Inherits BasePage

    Private Sql As String
    Private ds As DataSet
    Private Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ComposicaoDoEstoqueRN", "ACESSAR") Then
                BuscaUnidadeNegocio()
                BuscaDeposito()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Function Validar() As Boolean
        If ddlUnidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        Else
            Return True
        End If
    End Function

    Function RelatorioFiscal() As String
        Try
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Sql = "SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, Ped.Cliente, Ped.EndCliente, CliXP.Nome, " & vbCrLf &
                  "	NxI.Produto_Id as Produto, NotasFiscais.Operacao, NotasFiscais.SubOperacao, NotasFiscais.EntradaSaida_Id as EntradaSaida, 0 as Romaneio_Id, " & vbCrLf &
                  "	CONVERT(VARCHAR, NotasFiscais.Movimento, 103) AS Movimento, 0 as PesoLiquido, '' AS Processo, 0 AS Pesagem, " & vbCrLf &
                  "	NxI.EntradaSaida_Id + '-' + CAST(NxI.Nota_Id AS VARCHAR) + '-' + NxI.Serie_Id AS Nota_Id, " & vbCrLf &
                  "	CONVERT(VARCHAR, NotasFiscais.Movimento, 103) AS DataNota, " & vbCrLf &
                  "	NxI.QuantidadeFiscal as QuantidadeFisica, isnull(NxI.Deposito,'') AS DepositoNota, " & vbCrLf &
                  "	isnull(NxI.EndDeposito,0) AS EndDepositoNota, isnull(DNF.Nome,'') AS NomeDeposito " & vbCrLf &
                  "into #Temp " & vbCrLf &
                  "FROM NotasFiscais " & vbCrLf &
                  "	INNER JOIN NotasFiscaisxItens NxI " & vbCrLf &
                  "			 ON NxI.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf &
                  "			AND NxI.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                  "			AND NxI.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                  "			AND NxI.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                  "			AND NxI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                  "			AND NxI.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                  "			AND NxI.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                  "	INNER JOIN SubOperacoes " & vbCrLf &
                  "			 ON SubOperacoes.Operacao_Id     = NotasFiscais.Operacao " & vbCrLf &
                  "			AND SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao " & vbCrLf &
                  "	INNER JOIN Produtos Prod " & vbCrLf &
                  "			 ON Prod.Produto_Id = NxI.Produto_Id " & vbCrLf
            '"			AND Prod.Situacao   = 1 " & vbCrLf & _
            Sql &= "	INNER JOIN Clientes DNF " & vbCrLf &
                  "			 ON DNF.Cliente_Id  = NxI.Deposito " & vbCrLf &
                  "			AND DNF.Endereco_Id = NxI.EndDeposito " & vbCrLf &
                  "	LEFT JOIN Pedidos Ped " & vbCrLf &
                  "			 ON Ped.Empresa_Id    = NotasFiscais.Empresa_Id " & vbCrLf &
                  "			AND Ped.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id " & vbCrLf &
                  "			AND Ped.Pedido_Id     = NotasFiscais.Pedido " & vbCrLf &
                  "	LEFT JOIN Clientes CliXP " & vbCrLf &
                  "			 ON Ped.Cliente    = CliXP.Cliente_Id " & vbCrLf &
                  "			AND Ped.EndCliente = CliXP.Endereco_Id " & vbCrLf &
                  "WHERE (SubOperacoes.EstoqueFiscal   = 'S') " & vbCrLf &
                  "  AND (NotasFiscais.Situacao        = 1)" & vbCrLf &
                  "  AND (NotasFiscais.TipoDeDocumento = 1)" & vbCrLf &
                  "  AND (NotasFiscais.Empresa_Id      = '" & Empresa(0) & "') " & vbCrLf &
                  "  AND (NotasFiscais.EndEmpresa_Id   = " & Empresa(1) & ") " & vbCrLf &
                  "  AND (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

            If rdAtivo.Checked Then
                Sql &= " AND Prod.Situacao in(1) " & vbCrLf
            Else
                Sql &= " AND Prod.Situacao not in(1) " & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NxI.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            If ddlDeposito.SelectedIndex > 0 Then
                Dim Deposito() As String = ddlDeposito.SelectedValue.ToString.Split("-")
                Sql &= "			AND (CASE WHEN NxI.Deposito IS NULL " & vbCrLf &
                       "			          THEN NotasFiscais.Empresa_Id " & vbCrLf &
                       "			          ELSE NxI.Deposito END) = '" & Deposito(0) & "' " & vbCrLf
            End If

            If chkContraPartida.Checked Then
                Sql &= "insert into #Temp " & vbCrLf &
                    "SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, Ped.Cliente, Ped.EndCliente, CliXP.Nome, " & vbCrLf &
                    "	NxI.Produto_Id as Produto, NotasFiscais.Operacao, NotasFiscais.SubOperacao, NotasFiscais.EntradaSaida_Id as EntradaSaida, 0 as Romaneio_Id, " & vbCrLf &
                    "	CONVERT(VARCHAR, NotasFiscais.Movimento, 103) AS Movimento, 0 as PesoLiquido, '' AS Processo, 0 AS Pesagem, " & vbCrLf &
                    "	NxI.EntradaSaida_Id + '-' + CAST(NxI.Nota_Id AS VARCHAR) + '-' + NxI.Serie_Id AS Nota_Id, " & vbCrLf &
                    "	CONVERT(VARCHAR, NotasFiscais.Movimento, 103) AS DataNota, " & vbCrLf &
                    "	NxI.QuantidadeFiscal as QuantidadeFisica, isnull(NxI.Deposito,'') AS DepositoNota, " & vbCrLf &
                    "	isnull(NxI.EndDeposito,0) AS EndDepositoNota, isnull(DNF.Nome,'') AS NomeDeposito " & vbCrLf &
                    "FROM NotasFiscais " & vbCrLf &
                    "	INNER JOIN NotasFiscaisxItens NxI " & vbCrLf &
                    "			 ON NxI.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf &
                    "			AND NxI.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                    "			AND NxI.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                    "			AND NxI.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                    "			AND NxI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                    "			AND NxI.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                    "			AND NxI.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                    "	INNER JOIN SubOperacoes " & vbCrLf &
                    "			 ON SubOperacoes.Operacao_Id     = NotasFiscais.Operacao " & vbCrLf &
                    "			AND SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao " & vbCrLf &
                    "    INNER JOIN SubOperacoes SOD " & vbCrLf &
                    "       ON SOD.Operacao_Id     = SubOperacoes.OperacaoDestino  " & vbCrLf &
                    "      AND SOD.Suboperacoes_Id = SubOperacoes.SuboperacaoDestino  " & vbCrLf &
                    "	INNER JOIN Produtos Prod " & vbCrLf &
                    "			 ON Prod.Produto_Id = NxI.Produto_Id " & vbCrLf
                '"           AND Prod.Situacao   = 1 " & vbCrLf & _
                Sql &= "	INNER JOIN Clientes DNF " & vbCrLf &
                    "			 ON DNF.Cliente_Id  = NxI.Deposito " & vbCrLf &
                    "			AND DNF.Endereco_Id = NxI.EndDeposito " & vbCrLf &
                    "	LEFT JOIN Pedidos Ped " & vbCrLf &
                    "			 ON Ped.Empresa_Id    = NotasFiscais.Empresa_Id " & vbCrLf &
                    "			AND Ped.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id " & vbCrLf &
                    "			AND Ped.Pedido_Id     = NotasFiscais.Pedido " & vbCrLf &
                    "	LEFT JOIN Clientes CliXP " & vbCrLf &
                    "			 ON Ped.Cliente    = CliXP.Cliente_Id " & vbCrLf &
                    "			AND Ped.EndCliente = CliXP.Endereco_Id " & vbCrLf &
                    "WHERE (SubOperacoes.EstoqueFiscal   = 'S') " & vbCrLf &
                    "  AND (NotasFiscais.Situacao        = 1)" & vbCrLf &
                    "  AND (NotasFiscais.TipoDeDocumento = 1)" & vbCrLf &
                    "  AND (NotasFiscais.Empresa_Id      = '" & Empresa(0) & "') " & vbCrLf &
                    "  AND (NotasFiscais.EndEmpresa_Id   = " & Empresa(1) & ") " & vbCrLf &
                    "  AND (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

                If rdAtivo.Checked Then
                    Sql &= " AND Prod.Situacao in(1) " & vbCrLf
                Else
                    Sql &= " AND Prod.Situacao not in(1) " & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NxI.Produto_Id", "")
                    Sql &= " AND " & RetornoProdutos(0)
                End If

                If ddlDeposito.SelectedIndex > 0 Then
                    Dim Deposito() As String = ddlDeposito.SelectedValue.ToString.Split("-")
                    Sql &= "			AND ((CASE WHEN NxI.Deposito IS NULL " & vbCrLf &
                           "			          THEN NotasFiscais.Empresa_Id " & vbCrLf &
                           "			          ELSE NxI.Deposito END) = '" & Deposito(0) & "') " & vbCrLf
                End If

                Sql &= " " & vbCrLf
            End If

            Sql &= "select Empresa_id, EndEmpresa_Id, Pedido, Cliente, EndCliente, Nome, Produto, " & vbCrLf &
                    "       Operacao, SubOperacao, EntradaSaida, Romaneio_id, Movimento, PesoLiquido, " & vbCrLf &
                    "       Processo, Pesagem, Nota_Id, DataNota, QuantidadeFisica, DepositoNota, " & vbCrLf &
                    "       EndDepositoNota, NomeDeposito " & vbCrLf &
                    "from #Temp " & vbCrLf &
                    "ORDER BY Empresa_Id, EndEmpresa_Id, Produto, EntradaSaida, Movimento, Romaneio_Id " & vbCrLf
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return Sql
    End Function

    Function RelatorioFisico() As String
        Try
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Sql = "SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, Ped.Cliente, Ped.EndCliente, CliXP.Nome, " & vbCrLf &
            "	RO.Produto, RO.Operacao, RO.SubOperacao, RO.EntradaSaida, RO.Romaneio_Id, CONVERT(VARCHAR, RO.Movimento, 103) AS Movimento, RO.PesoLiquido, " & vbCrLf &
            "	isnull(RO.Processo,'') AS Processo, isnull(RxP.Pesagem_Id,0) AS Pesagem, " & vbCrLf &
            "	CASE " & vbCrLf &
            "		WHEN ISNULL(NxI.Nota_Id,0) > 0 " & vbCrLf &
            "			THEN NxI.EntradaSaida_Id + '-' + CAST(NxI.Nota_Id AS VARCHAR) + '-' + NxI.Serie_Id " & vbCrLf &
            "			ELSE '0' " & vbCrLf &
            "		END AS Nota_Id, " & vbCrLf &
            "	CASE " & vbCrLf &
            "		WHEN ISNULL(NotasFiscais.Movimento,'1900') = '1900' " & vbCrLf &
            "			THEN '' " & vbCrLf &
            "			ELSE CONVERT(VARCHAR, NotasFiscais.Movimento, 103) " & vbCrLf &
            "		END AS DataNota, " & vbCrLf &
            "	NxI.QuantidadeFisica, isnull(NxI.Deposito,'') AS DepositoNota, " & vbCrLf &
            "	isnull(NxI.EndDeposito,0) AS EndDepositoNota, isnull(DNF.Nome,'') AS NomeDeposito " & vbCrLf &
            "into #Temp " & vbCrLf &
            "FROM SubOperacoes " & vbCrLf &
            "	INNER JOIN Romaneios RO " & vbCrLf &
            "			 ON SubOperacoes.Operacao_Id     = RO.Operacao " & vbCrLf &
            "			AND SubOperacoes.SubOperacoes_Id = RO.SubOperacao " & vbCrLf &
            "	INNER JOIN Produtos Prod " & vbCrLf &
            "			 ON Prod.Produto_Id = RO.Produto " & vbCrLf
            '"			AND Prod.Situacao   = 1 " & vbCrLf & _
            Sql &= "	LEFT JOIN NotasFiscaisXRomaneios NxR " & vbCrLf &
            "			 ON RO.Empresa_Id    = NxR.Empresa_Id " & vbCrLf &
            "			AND RO.EndEmpresa_Id = NxR.EndEmpresa_Id " & vbCrLf &
            "			AND RO.Romaneio_Id   = NxR.Romaneio_Id " & vbCrLf &
            "	LEFT JOIN NotasFiscaisxItens NxI " & vbCrLf &
            "			 ON NxR.Empresa_Id      = NxI.Empresa_Id " & vbCrLf &
            "			AND NxR.EndEmpresa_Id   = NxI.EndEmpresa_Id  " & vbCrLf &
            "			AND NxR.Cliente_Id      = NxI.Cliente_Id  " & vbCrLf &
            "			AND NxR.EndCliente_Id   = NxI.EndCliente_Id  " & vbCrLf &
            "			AND NxR.EntradaSaida_Id = NxI.EntradaSaida_Id  " & vbCrLf &
            "			AND NxR.Serie_Id        = NxI.Serie_Id  " & vbCrLf &
            "			AND NxR.Nota_Id         = NxI.Nota_Id  " & vbCrLf &
            "	LEFT JOIN NotasFiscais " & vbCrLf &
            "			 ON NxI.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
            "			AND NxI.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
            "			AND NxI.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
            "			AND NxI.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
            "			AND NxI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
            "			AND NxI.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
            "			AND NxI.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
            "	LEFT JOIN Clientes DNF " & vbCrLf &
            "			 ON NxI.Deposito    = DNF.Cliente_Id " & vbCrLf &
            "			AND NxI.EndDeposito = DNF.Endereco_Id " & vbCrLf &
            "	LEFT JOIN Pedidos Ped " & vbCrLf &
            "			 ON Ped.Empresa_Id    = RO.Empresa_Id " & vbCrLf &
            "			AND Ped.EndEmpresa_Id = RO.EndEmpresa_Id " & vbCrLf &
            "			AND Ped.Pedido_Id     = RO.Pedido " & vbCrLf &
            "	LEFT JOIN Clientes CliXP " & vbCrLf &
            "			 ON Ped.Cliente    = CliXP.Cliente_Id " & vbCrLf &
            "			AND Ped.EndCliente = CliXP.Endereco_Id " & vbCrLf &
            "	LEFT JOIN RomaneiosXPesagens RxP " & vbCrLf &
            "			 ON RO.Empresa_Id    = RxP.Empresa_Id " & vbCrLf &
            "			AND RO.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf &
            "			AND RO.Romaneio_Id   = RxP.Romaneio_Id " & vbCrLf &
            "WHERE (SubOperacoes.EstoqueFisico = 'S') " & vbCrLf &
            "  AND (RO.Empresa_Id = '" & Empresa(0) & "') " & vbCrLf &
            "  AND (RO.EndEmpresa_Id = " & Empresa(1) & ") " & vbCrLf &
            "  AND (RO.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

            If rdAtivo.Checked Then
                Sql &= " AND Prod.Situacao in(1) " & vbCrLf
            Else
                Sql &= " AND Prod.Situacao not in(1) " & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "RO.Produto", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            If ddlDeposito.SelectedIndex > 0 Then
                Dim Deposito() As String = ddlDeposito.SelectedValue.ToString.Split("-")
                Sql &= "			AND (CASE WHEN NxI.Deposito IS NULL " & vbCrLf &
                       "			          THEN RO.Deposito " & vbCrLf &
                       "			          ELSE NxI.Deposito END) = '" & Deposito(0) & "' " & vbCrLf
            End If

            If chkContraPartida.Checked Then
                Sql &= "insert into #Temp " & vbCrLf &
                "   SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, Ped.Cliente, Ped.EndCliente, CliXP.Nome, " & vbCrLf &
                "RO.Produto, SOD.Operacao_Id, SOD.Suboperacoes_Id, SOD.EntradaSaida, RO.Romaneio_Id, CONVERT(VARCHAR, RO.Movimento, 103) AS Movimento, RO.PesoLiquido, " & vbCrLf &
                "	isnull(RO.Processo,'') AS Processo, isnull(RxP.Pesagem_Id,0) AS Pesagem, " & vbCrLf &
                "	CASE " & vbCrLf &
                "		WHEN ISNULL(NxI.Nota_Id,0) > 0 " & vbCrLf &
                "			THEN NxI.EntradaSaida_Id + '-' + CAST(NxI.Nota_Id AS VARCHAR) + '-' + NxI.Serie_Id " & vbCrLf &
                "			ELSE '0' " & vbCrLf &
                "		END AS Nota_Id, " & vbCrLf &
                "	CASE " & vbCrLf &
                "		WHEN ISNULL(N.Movimento,'1900') = '1900' " & vbCrLf &
                "			THEN '' " & vbCrLf &
                "			ELSE CONVERT(VARCHAR, N.Movimento, 103) " & vbCrLf &
                "		END AS DataNota, " & vbCrLf &
                "	NxI.QuantidadeFisica, isnull(N.Destino,'') AS DepositoNota, " & vbCrLf &
                "	isnull(N.EndDestino,0) AS EndDepositoNota, isnull(DNF.Nome,'') AS NomeDeposito " & vbCrLf &
                "     FROM SubOperacoes " & vbCrLf &
                "    INNER JOIN Romaneios RO " & vbCrLf &
                "       ON SubOperacoes.Operacao_Id     = RO.Operacao " & vbCrLf &
                "      AND SubOperacoes.SubOperacoes_Id = RO.SubOperacao " & vbCrLf &
                "    INNER JOIN SubOperacoes SOD " & vbCrLf &
                "       ON SOD.Operacao_Id     = SubOperacoes.OperacaoDestino  " & vbCrLf &
                "      AND SOD.Suboperacoes_Id = SubOperacoes.SuboperacaoDestino  " & vbCrLf &
                "    INNER JOIN NotasFiscaisXRomaneios NxR " & vbCrLf &
                "       ON RO.Empresa_Id    = NxR.Empresa_Id " & vbCrLf &
                "      AND RO.EndEmpresa_Id = NxR.EndEmpresa_Id " & vbCrLf &
                "      AND RO.Romaneio_Id   = NxR.Romaneio_Id " & vbCrLf &
                "	INNER JOIN NotasFiscaisxItens NxI " & vbCrLf &
                "			 ON NxR.Empresa_Id      = NxI.Empresa_Id " & vbCrLf &
                "			AND NxR.EndEmpresa_Id   = NxI.EndEmpresa_Id  " & vbCrLf &
                "			AND NxR.Cliente_Id      = NxI.Cliente_Id  " & vbCrLf &
                "			AND NxR.EndCliente_Id   = NxI.EndCliente_Id  " & vbCrLf &
                "			AND NxR.EntradaSaida_Id = NxI.EntradaSaida_Id  " & vbCrLf &
                "			AND NxR.Serie_Id        = NxI.Serie_Id  " & vbCrLf &
                "			AND NxR.Nota_Id         = NxI.Nota_Id  " & vbCrLf &
                "	INNER JOIN NotasFiscais N " & vbCrLf &
                "			 ON NxI.Empresa_Id      = N.Empresa_Id  " & vbCrLf &
                "			AND NxI.EndEmpresa_Id   = N.EndEmpresa_Id  " & vbCrLf &
                "			AND NxI.Cliente_Id      = N.Cliente_Id  " & vbCrLf &
                "			AND NxI.EndCliente_Id   = N.EndCliente_Id  " & vbCrLf &
                "			AND NxI.EntradaSaida_Id = N.EntradaSaida_Id  " & vbCrLf &
                "			AND NxI.Serie_Id        = N.Serie_Id  " & vbCrLf &
                "			AND NxI.Nota_Id         = N.Nota_Id  " & vbCrLf &
                "    INNER JOIN Produtos Prod  " & vbCrLf &
                "           ON Prod.Produto_Id = NxI.Produto_Id " & vbCrLf
                '"			AND Prod.Situacao  = 1  " & vbCrLf & _
                Sql &= "	LEFT JOIN Clientes DNF " & vbCrLf &
                "			 ON N.Destino    = DNF.Cliente_Id " & vbCrLf &
                "			AND N.EndDestino = DNF.Endereco_Id " & vbCrLf &
                "	LEFT JOIN Pedidos Ped " & vbCrLf &
                "			 ON Ped.Empresa_Id    = RO.Empresa_Id " & vbCrLf &
                "			AND Ped.EndEmpresa_Id = RO.EndEmpresa_Id " & vbCrLf &
                "			AND Ped.Pedido_Id     = RO.Pedido " & vbCrLf &
                "	LEFT JOIN Clientes CliXP " & vbCrLf &
                "			 ON Ped.Cliente    = CliXP.Cliente_Id " & vbCrLf &
                "			AND Ped.EndCliente = CliXP.Endereco_Id " & vbCrLf &
                "	LEFT JOIN RomaneiosXPesagens RxP " & vbCrLf &
                "			 ON RO.Empresa_Id    = RxP.Empresa_Id " & vbCrLf &
                "			AND RO.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf &
                "			AND RO.Romaneio_Id   = RxP.Romaneio_Id " & vbCrLf &
                "    WHERE (SOD.EstoqueFisico = 'S') " & vbCrLf &
                "      AND (N.Situacao        = 1) " & vbCrLf &
                "      AND (N.TipoDeDocumento = 1) " & vbCrLf &
                "      AND (SubOperacoes.Deposito = 'S') " & vbCrLf &
                "      AND (RO.Empresa_Id = '" & Empresa(0) & "') " & vbCrLf &
                "      AND (RO.EndEmpresa_Id = " & Empresa(1) & ") " & vbCrLf &
                "      AND (RO.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

                If rdAtivo.Checked Then
                    Sql &= " AND Prod.Situacao in(1) " & vbCrLf
                Else
                    Sql &= " AND Prod.Situacao not in(1) " & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "RO.Produto", "")
                    Sql &= " AND " & RetornoProdutos(0)
                End If

                If ddlDeposito.SelectedIndex > 0 Then
                    Dim Deposito() As String = ddlDeposito.SelectedValue.ToString.Split("-")
                    Sql &= "      AND (N.Destino = '" & Deposito(0) & "') " & vbCrLf
                End If

                Sql &= " " & vbCrLf
            End If

            Sql &= "select Empresa_id, EndEmpresa_Id, Pedido, Cliente, EndCliente, Nome, Produto, " & vbCrLf &
                    "       Operacao, SubOperacao, EntradaSaida, Romaneio_id, Movimento, PesoLiquido, " & vbCrLf &
                    "       Processo, Pesagem, Nota_Id, DataNota, QuantidadeFisica, DepositoNota, " & vbCrLf &
                    "       EndDepositoNota, NomeDeposito " & vbCrLf &
                    "from #Temp " & vbCrLf &
                    "ORDER BY Empresa_Id, EndEmpresa_Id, Produto, EntradaSaida, Movimento, Romaneio_Id " & vbCrLf

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return Sql
    End Function

    Private Sub BuscaUnidadeNegocio()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub BuscaDeposito()
        ddl.Carregar(ddlDeposito, CarregarDDL.Tabela.ClientesXTipos, "3", True)
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

        ddlDeposito.SelectedIndex = 0

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Parametros:" & vbCrLf & "Unidade: " & ddlUnidade.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= " à: " & txtDataFinal.Text & vbCrLf
        End If

        param &= IIf(rdFisico.Checked, "Estoque: : Físico", "Estoque: : Fiscal")

        Return param
    End Function

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ComposicaoDoEstoqueRN", "RELATORIO") Then
                If Validar() Then
                    Sql = IIf(rdFisico.Checked, RelatorioFisico(), RelatorioFiscal())
                    ds = Banco.ConsultaDataSet(Sql, "ComposicaoDoEstoqueRN")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Titulo", "Posição de Estoques Romaneios/Nota Fiscal.")
                    parameters.Add("Parametros", getParam())

                    Funcoes.BindReport(Me.Page, ds, "Cr_ComposicaoDoEstoqueRN", eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ComposicaoDoEstoqueRN")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class