Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class ComposicaoDoEstoque
    Inherits BasePage

    Private Sql As String
    Private ds As DataSet
    Private Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ComposicaoDoEstoque", "ACESSAR") Then
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
            MsgBox(Me.Page, "Unidade de negócio é obrigatório")
            Return False
        ElseIf ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa é obrigatório")
            Return False
        Else
            Return True
        End If
    End Function

    Function RelatorioFiscal() As String
        Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

        Sql = "SELECT Consulta.Empresa, Consulta.EndEmpresa, " & vbCrLf & _
        "       Consulta.Deposito, Consulta.EndDeposito, " & vbCrLf & _
        "       Depositos.Nome AS NomeDoDeposito, Depositos.Cidade AS CidadeDoDeposito, " & vbCrLf & _
        "       Depositos.Estado AS EstadoDoDeposito, Depositos.Reduzido AS ReduzidoDoDeposito, " & vbCrLf & _
        "       ProdutosConsulta.Grupo, GruposDeEstoques.Descricao as NomeDoGrupo, " & vbCrLf & _
        "	   Consulta.Produto, ProdutosConsulta.Nome as NomeDoProduto, " & vbCrLf & _
        "       Consulta.Movimento, " & vbCrLf & _
        "       SUM(Consulta.Entradas) AS Entradas, SUM(Consulta.Saidas) AS Saidas, " & vbCrLf & _
        "       SUM(Consulta.Producao) AS Producao, SUM(Consulta.Consumo) AS Consumo " & vbCrLf & _
        "     INTO #Produto" & vbCrLf & _
        "  FROM ( " & vbCrLf & _
        "--------------PRODUCAO " & vbCrLf & _
        "        Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto_Id as Produto, " & vbCrLf & _
        "               0 AS Entradas, " & vbCrLf & _
        "               0 AS Saidas, " & vbCrLf & _
        "               isnull(Entradas,0) As Producao, " & vbCrLf & _
        "               isnull(Saidas,0) As Consumo, " & vbCrLf & _
        "               Movimento " & vbCrLf & _
        "         From ( " & vbCrLf & _
        "               SELECT Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id As Deposito, Producao.EndDeposito_Id as EndDeposito, Producao.Produto_Id, Producao.Movimento_Id as Movimento, " & vbCrLf & _
        "                      ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Entradas END), 0) AS Entradas, " & vbCrLf & _
        "                      ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Saidas END), 0) AS Saidas " & vbCrLf & _
        "                 FROM SubOperacoes  " & vbCrLf & _
        "                INNER JOIN Producao " & vbCrLf & _
        "                   ON SubOperacoes.Operacao_Id     = Producao.Operacao_Id " & vbCrLf & _
        "                  AND SubOperacoes.SubOperacoes_Id = Producao.SubOperacao_Id " & vbCrLf & _
        "                INNER JOIN Produtos " & vbCrLf & _
        "                   ON Produtos.Produto_Id = Producao.Produto_Id " & vbCrLf
        '"                  AND Produtos.Situacao   = 1 " & vbCrLf & _
        Sql &= "                WHERE SubOperacoes.EstoqueFiscal   = 'S' " & vbCrLf & _
        "                  AND Producao.FisicoFiscal_Id     = 2 " & vbCrLf & _
        "                  And Producao.Empresa_Id          = '" & Empresa(0) & "' " & vbCrLf & _
        "                  AND Producao.EndEmpresa_Id       = " & Empresa(1) & vbCrLf & _
        "                  AND Producao.Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
        "                GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Movimento_Id " & vbCrLf & _
        "               ) as ConsultaProducao " & vbCrLf & _
        "          Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto_Id, Movimento,Entradas,Saidas " & vbCrLf & _
        "------------------------- FIM PRODUCAO " & vbCrLf & _
        "          UNION " & vbCrLf & _
        "------------------NOTAS FISCAIS " & vbCrLf & _
        "         Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto, " & vbCrLf & _
        "                Isnull(Entradas,0) AS Entradas, " & vbCrLf & _
        "           isnull(Saidas,0) As Saidas, " & vbCrLf & _
        "                0 as Producao, " & vbCrLf & _
        "                0 as Consumo, " & vbCrLf & _
        "                Movimento " & vbCrLf & _
        "           From ( " & vbCrLf & _
        "                 SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, isnull(nfxi.Deposito,NotasFiscais.Empresa_Id) as Deposito, " & vbCrLf & _
        "                        isnull(nfxi.EndDeposito,NotasFiscais.EndEmpresa_Id) as EndDeposito, nfxi.Produto_Id as Produto, NotasFiscais.Movimento, " & vbCrLf & _
        "                        isnull(SUM(CASE  WHEN NotasFiscais.EntradaSaida_Id = 'E' THEN nfxi.QuantidadeFiscal  END),0) AS Entradas, " & vbCrLf & _
        "                        isnull(SUM(CASE  WHEN NotasFiscais.EntradaSaida_Id = 'S' THEN nfxi.QuantidadeFiscal  END),0) AS Saidas " & vbCrLf & _
        "                   FROM NotasFiscais  " & vbCrLf & _
        "                  INNER JOIN NotasFiscaisxItens nfxi " & vbCrLf & _
        "                     ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf & _
        "                    AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
        "                    AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
        "                    AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
        "                    AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
        "                    AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
        "                    AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
        "                  INNER JOIN SubOperacoes " & vbCrLf & _
        "                     ON SubOperacoes.Operacao_Id     = nfxi.Operacao " & vbCrLf & _
        "                    AND SubOperacoes.SubOperacoes_Id = nfxi.SubOperacao " & vbCrLf & _
        "                  INNER JOIN Produtos Prod  " & vbCrLf & _
        "                     ON Prod.Produto_Id = nfxi.Produto_Id  " & vbCrLf
        '"                    AND Prod.Situacao   = 1 " & vbCrLf & _
        Sql &= "                  WHERE (SubOperacoes.EstoqueFiscal = 'S') " & vbCrLf &
        "                    AND NotasFiscais.Situacao        = 1 " & vbCrLf &
        "                    AND NotasFiscais.TipoDeDocumento = 1 " & vbCrLf &
        "                    AND (NotasFiscais.Empresa_Id = '" & Empresa(0) & "') " & vbCrLf &
        "                    AND (NotasFiscais.EndEmpresa_Id = " & Empresa(1) & ") " & vbCrLf &
        "                    AND (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
        "                 GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, nfxi.Deposito, nfxi.EndDeposito, nfxi.Produto_Id, NotasFiscais.Movimento " & vbCrLf

        If chkContraPartida.Checked Then
            Sql &= "    UNION ALL " & vbCrLf & _
            "------CONTRA PARTIDA NOTAS FISCAIS " & vbCrLf & _
            "                 SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, isnull(nfxi.Deposito,NotasFiscais.Empresa_Id) as Deposito, " & vbCrLf & _
            "                        isnull(nfxi.EndDeposito,NotasFiscais.EndEmpresa_Id) as EndDeposito, nfxi.Produto_Id as Produto, NotasFiscais.Movimento, " & vbCrLf & _
            "                        isnull(SUM(CASE  WHEN NotasFiscais.EntradaSaida_Id = 'E' THEN nfxi.QuantidadeFiscal  END),0) AS Entradas, " & vbCrLf & _
            "                        isnull(SUM(CASE  WHEN NotasFiscais.EntradaSaida_Id = 'S' THEN nfxi.QuantidadeFiscal  END),0) AS Saidas " & vbCrLf & _
            "                   FROM NotasFiscais  " & vbCrLf & _
            "                  INNER JOIN NotasFiscaisxItens nfxi " & vbCrLf & _
            "                     ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf & _
            "                    AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
            "                    AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
            "                    AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
            "                    AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
            "                    AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
            "                    AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
            "                  INNER JOIN SubOperacoes " & vbCrLf & _
            "                     ON SubOperacoes.Operacao_Id     = nfxi.Operacao " & vbCrLf & _
            "                    AND SubOperacoes.SubOperacoes_Id = nfxi.SubOperacao " & vbCrLf & _
            "                  INNER JOIN SubOperacoes SOD " & vbCrLf & _
            "                     ON SOD.Operacao_Id     = SubOperacoes.OperacaoDestino " & vbCrLf & _
            "                    AND SOD.SubOperacoes_Id = SubOperacoes.SuboperacaoDestino " & vbCrLf & _
            "                  INNER JOIN Produtos Prod  " & vbCrLf & _
            "                     ON Prod.Produto_Id = nfxi.Produto_Id  " & vbCrLf
            '"                    AND Prod.Situacao   = 1 " & vbCrLf & _
            Sql &= "                  WHERE (SubOperacoes.EstoqueFiscal = 'S') " & vbCrLf &
            "                    AND NotasFiscais.Situacao        = 1 " & vbCrLf &
            "                    AND NotasFiscais.TipoDeDocumento = 1 " & vbCrLf &
            "                    AND (NotasFiscais.Empresa_Id     = '" & Empresa(0) & "') " & vbCrLf &
            "                    AND (NotasFiscais.EndEmpresa_Id  = " & Empresa(1) & ") " & vbCrLf &
            "                    AND (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "                 GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, nfxi.Deposito, nfxi.EndDeposito, nfxi.Produto_Id, NotasFiscais.Movimento " & vbCrLf
        End If

        Sql &= "                 ) as ConsultaRomaneios " & vbCrLf & _
        "            Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto,Movimento,Entradas,Saidas " & vbCrLf & _
        "--------------FIM NOTAS FISCAIS " & vbCrLf & _
        "           ) AS Consulta " & vbCrLf & _
        "      INNER JOIN Clientes AS Empresa ON Consulta.Empresa = Empresa.Cliente_Id AND Consulta.EndEmpresa = Empresa.Endereco_Id " & vbCrLf & _
        "      INNER JOIN Clientes AS Depositos ON Consulta.Deposito = Depositos.Cliente_Id AND Consulta.EndDeposito = Depositos.Endereco_Id " & vbCrLf & _
        "      INNER JOIN Produtos AS ProdutosConsulta ON Consulta.Produto = ProdutosConsulta.Produto_Id " & vbCrLf & _
        "      INNER JOIN GruposDeEstoques ON ProdutosConsulta.Grupo = GruposDeEstoques.Grupo_Id " & vbCrLf

        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList
            RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Consulta.Produto", "")
            Sql &= " AND " & RetornoProdutos(0)
        End If

        Sql &= "        and (Consulta.Entradas <> 0 " & vbCrLf & _
        "         or Consulta.Producao <> 0 " & vbCrLf & _
        "         or Consulta.Consumo <> 0 " & vbCrLf & _
        "         or Consulta.Saidas <> 0) " & vbCrLf

        If ddlDeposito.SelectedIndex > 0 Then
            Sql &= "           and  (Consulta.Deposito = '" & ddlDeposito.SelectedValue & "') " & vbCrLf
        End If

        If rdAtivo.Checked Then
            Sql &= " AND ProdutosConsulta.Situacao in(1) " & vbCrLf
        Else
            Sql &= " AND ProdutosConsulta.Situacao not in(1) " & vbCrLf
        End If

        Sql &= " " & vbCrLf & _
        "GROUP BY Consulta.Empresa, Consulta.EndEmpresa, " & vbCrLf & _
        "       Consulta.Deposito, Consulta.EndDeposito, " & vbCrLf & _
        "       Depositos.Nome, Depositos.Cidade, " & vbCrLf & _
        "       Depositos.Estado, Depositos.Reduzido, " & vbCrLf & _
        "       ProdutosConsulta.Grupo, GruposDeEstoques.Descricao, " & vbCrLf & _
        "       Consulta.Produto, ProdutosConsulta.Nome, " & vbCrLf & _
        "       Consulta.Movimento " & vbCrLf & _
        " " & vbCrLf & _
        " " & vbCrLf & _
        "select Deposito, EndDeposito, NomeDoDeposito, CidadeDoDeposito, EstadoDoDeposito, ReduzidoDoDeposito, " & vbCrLf & _
        "		Grupo, NomeDoGrupo, Produto, NomeDoProduto, Movimento, Entradas, Saidas, Producao, Consumo " & vbCrLf & _
        "from #Produto" & vbCrLf & _
        "ORDER BY Deposito, EndDeposito, Produto, Movimento" & vbCrLf & _
        " " & vbCrLf & _
        " " & vbCrLf & _
        "select Produto, NomeDoProduto, sum(Entradas) as Entradas, sum(Producao) As Producao, sum(Consumo) as Consumo, sum(Saidas) as Saidas " & vbCrLf & _
        "from #Produto " & vbCrLf & _
        "group by Produto, NomeDoProduto " & vbCrLf & _
        "order by NomeDoProduto"

        Return Sql
    End Function

    Function RelatorioFisico() As String
        Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

        Sql = "SELECT Consulta.Empresa, Consulta.EndEmpresa, " & vbCrLf & _
        "       Consulta.Deposito, Consulta.EndDeposito, " & vbCrLf & _
        "       Depositos.Nome AS NomeDoDeposito, Depositos.Cidade AS CidadeDoDeposito, " & vbCrLf & _
        "       Depositos.Estado AS EstadoDoDeposito, Depositos.Reduzido AS ReduzidoDoDeposito, " & vbCrLf & _
        "       ProdutosConsulta.Grupo, GruposDeEstoques.Descricao as NomeDoGrupo, " & vbCrLf & _
        "	   Consulta.Produto, ProdutosConsulta.Nome as NomeDoProduto, " & vbCrLf & _
        "       Consulta.Movimento, " & vbCrLf & _
        "       SUM(Consulta.Entradas) AS Entradas, SUM(Consulta.Saidas) AS Saidas, " & vbCrLf & _
        "       SUM(Consulta.Producao) AS Producao, SUM(Consulta.Consumo) AS Consumo " & vbCrLf & _
        "     INTO #Produto" & vbCrLf & _
        "  FROM ( " & vbCrLf & _
        "--------------PRODUCAO " & vbCrLf & _
        "        Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto_Id as Produto, " & vbCrLf & _
        "               0 AS Entradas, " & vbCrLf & _
        "               0 AS Saidas, " & vbCrLf & _
        "               isnull(Entradas,0) As Producao, " & vbCrLf & _
        "               isnull(Saidas,0) As Consumo, " & vbCrLf & _
        "               Movimento " & vbCrLf & _
        "         From ( " & vbCrLf & _
        "               SELECT Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id As Deposito, Producao.EndDeposito_Id as EndDeposito, Producao.Produto_Id, Producao.Movimento_Id as Movimento, " & vbCrLf & _
        "                      ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Entradas END), 0) AS Entradas, " & vbCrLf & _
        "                      ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Saidas END), 0) AS Saidas " & vbCrLf & _
        "                 FROM SubOperacoes  " & vbCrLf & _
        "                INNER JOIN Producao " & vbCrLf & _
        "                   ON SubOperacoes.Operacao_Id     = Producao.Operacao_Id " & vbCrLf & _
        "                  AND SubOperacoes.SubOperacoes_Id = Producao.SubOperacao_Id " & vbCrLf & _
        "                INNER JOIN Produtos " & vbCrLf & _
        "                   ON Produtos.Produto_Id = Producao.Produto_Id " & vbCrLf
        '"                  AND Produtos.Situacao   = 1 " & vbCrLf & _
        Sql &= "                WHERE SubOperacoes.EstoqueFisico   = 'S' " & vbCrLf & _
        "                  AND Producao.FisicoFiscal_Id     = 1 " & vbCrLf & _
        "                  And Producao.Empresa_Id          = '" & Empresa(0) & "' " & vbCrLf & _
        "                  AND Producao.EndEmpresa_Id       = " & Empresa(1) & vbCrLf & _
        "                  AND Producao.Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
        "                GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Movimento_Id " & vbCrLf & _
        "               ) as ConsultaProducao " & vbCrLf & _
        "          Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto_Id, Movimento,Entradas,Saidas " & vbCrLf & _
        "------------------------- FIM PRODUCAO " & vbCrLf & _
        "          UNION " & vbCrLf & _
        "------------------ROMANEIOS " & vbCrLf & _
        "         Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto, " & vbCrLf & _
        "                Isnull(Entradas,0) AS Entradas, " & vbCrLf & _
        "           isnull(Saidas,0) As Saidas, " & vbCrLf & _
        "                0 as Producao, " & vbCrLf & _
        "                0 as Consumo, " & vbCrLf & _
        "                Movimento " & vbCrLf & _
        "           From ( " & vbCrLf & _
        "                 SELECT Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, nfxi.Deposito, nfxi.EndDeposito, Romaneios.Produto, " & vbCrLf & _
        "                        Romaneios.Movimento, " & vbCrLf & _
        "               ISNULL(SUM(CASE WHEN SubOperacoes.EntradaSaida = 'E' THEN CASE WHEN Prod.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Entradas, " & vbCrLf & _
        "               ISNULL(SUM(CASE WHEN SubOperacoes.EntradaSaida = 'S' THEN CASE WHEN Prod.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Saidas " & vbCrLf & _
        "                   FROM SubOperacoes  " & vbCrLf & _
        "                  INNER JOIN Romaneios " & vbCrLf & _
        "                     ON SubOperacoes.Operacao_Id     = Romaneios.Operacao " & vbCrLf & _
        "                    AND SubOperacoes.SubOperacoes_Id = Romaneios.SubOperacao " & vbCrLf & _
        "                  INNER JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf & _
        "                     ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
        "                    AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
        "                    AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
        "                  INNER JOIN NotasFiscais " & vbCrLf & _
        "                     ON nfxr.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
        "                    AND nfxr.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
        "                    AND nfxr.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
        "                    AND nfxr.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
        "                    AND nfxr.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
        "                    AND nfxr.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
        "                    AND nfxr.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
        "                  INNER JOIN NotasFiscaisxItens nfxi " & vbCrLf & _
        "                     ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf & _
        "                    AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
        "                    AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
        "                    AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
        "                    AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
        "                    AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
        "                    AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
        "                  INNER JOIN Produtos Prod  " & vbCrLf & _
        "                     ON Prod.Produto_Id = nfxi.Produto_Id  " & vbCrLf
        '"                    AND Prod.Situacao   = 1  " & vbCrLf & _
        Sql &= "                  WHERE (SubOperacoes.EstoqueFisico = 'S') " & vbCrLf &
        "                    AND NotasFiscais.Situacao        = 1 " & vbCrLf &
        "                    AND NotasFiscais.TipoDeDocumento = 1 " & vbCrLf &
        "                    AND (Romaneios.Empresa_Id        = '" & Empresa(0) & "') " & vbCrLf &
        "                    AND (Romaneios.EndEmpresa_Id     = " & Empresa(1) & ") " & vbCrLf &
        "                    And (Romaneios.Deposito <> '') " & vbCrLf &
        "                    AND (Romaneios.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
        "                 GROUP BY Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, nfxi.Deposito, nfxi.EndDeposito, Romaneios.Produto, Romaneios.Movimento " & vbCrLf

        If chkContraPartida.Checked Then
            Sql &= "    UNION ALL " & vbCrLf & _
            "------CONTRA PARTIDA ROMANEIOS " & vbCrLf & _
            "   SELECT Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id,  NotasFiscais.Destino, NotasFiscais.EndDestino, Romaneios.Produto, " & vbCrLf & _
            "          Romaneios.Movimento, " & vbCrLf & _
            "          ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN CASE WHEN Prod.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Entradas, " & vbCrLf & _
            "          ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'S' THEN CASE WHEN Prod.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Saidas " & vbCrLf & _
            "     FROM SubOperacoes " & vbCrLf & _
            "    INNER JOIN Romaneios " & vbCrLf & _
            "       ON SubOperacoes.Operacao_Id     = Romaneios.Operacao " & vbCrLf & _
            "      AND SubOperacoes.SubOperacoes_Id = Romaneios.SubOperacao " & vbCrLf & _
            "    INNER JOIN SubOperacoes SOD " & vbCrLf & _
            "       ON SOD.Operacao_Id     = SubOperacoes.OperacaoDestino  " & vbCrLf & _
            "      AND SOD.Suboperacoes_Id = SubOperacoes.SuboperacaoDestino " & vbCrLf & _
            "    INNER JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf & _
            "       ON Romaneios.Empresa_Id = nfxr.Empresa_Id " & vbCrLf & _
            "      AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
            "      AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
            "    INNER JOIN NotasFiscais " & vbCrLf & _
            "       ON nfxr.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
            "      AND nfxr.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
            "      AND nfxr.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
            "      AND nfxr.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
            "      AND nfxr.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
            "      AND nfxr.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
            "      AND nfxr.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
            "    INNER JOIN NotasFiscaisxItens nfxi " & vbCrLf & _
            "       ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
            "      AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
            "      AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
            "      AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
            "      AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
            "      AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
            "      AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
            "    INNER JOIN Produtos Prod  " & vbCrLf & _
            "       ON Prod.Produto_Id = nfxi.Produto_Id " & vbCrLf
            '"      AND Prod.Situacao   = 1  " & vbCrLf & _
            Sql &= "    WHERE (SOD.EstoqueFisico = 'S') " & vbCrLf &
            "      AND NotasFiscais.Situacao        = 1 " & vbCrLf &
            "      AND NotasFiscais.TipoDeDocumento = 1 " & vbCrLf &
            "      AND SubOperacoes.Deposito        = 'S' " & vbCrLf &
            "      AND (Romaneios.Empresa_Id = '" & Empresa(0) & "') " & vbCrLf &
            "      AND (Romaneios.EndEmpresa_Id = " & Empresa(1) & ") " & vbCrLf &
            "      And (Romaneios.Deposito <> '') " & vbCrLf &
            "      AND (Romaneios.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
            "     GROUP BY Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, NotasFiscais.Destino, NotasFiscais.EndDestino, Romaneios.Produto, Romaneios.Movimento " & vbCrLf
        End If

        Sql &= "                 ) as ConsultaRomaneios " & vbCrLf & _
        "            Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto,Movimento,Entradas,Saidas " & vbCrLf & _
        "--------------FIM ROMANEIOS " & vbCrLf & _
        "           ) AS Consulta " & vbCrLf & _
        "      INNER JOIN Clientes AS Empresa ON Consulta.Empresa = Empresa.Cliente_Id AND Consulta.EndEmpresa = Empresa.Endereco_Id " & vbCrLf & _
        "      INNER JOIN Clientes AS Depositos ON Consulta.Deposito = Depositos.Cliente_Id AND Consulta.EndDeposito = Depositos.Endereco_Id " & vbCrLf & _
        "      INNER JOIN Produtos AS ProdutosConsulta ON Consulta.Produto = ProdutosConsulta.Produto_Id " & vbCrLf & _
        "      INNER JOIN GruposDeEstoques ON ProdutosConsulta.Grupo = GruposDeEstoques.Grupo_Id " & vbCrLf

        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList
            RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Consulta.Produto", "")
            Sql &= " AND " & RetornoProdutos(0)
        End If

        Sql &= "        and (Consulta.Entradas <> 0 " & vbCrLf & _
        "         or Consulta.Producao <> 0 " & vbCrLf & _
        "         or Consulta.Consumo <> 0 " & vbCrLf & _
        "         or Consulta.Saidas <> 0) " & vbCrLf

        If ddlDeposito.SelectedIndex > 0 Then
            Dim Deposito() As String = ddlDeposito.SelectedValue.ToString.Split("-")
            Sql &= "           and  (Consulta.Deposito = '" & Deposito(0) & "') " & vbCrLf
        End If

        If rdAtivo.Checked Then
            Sql &= " AND ProdutosConsulta.Situacao in(1) " & vbCrLf
        Else
            Sql &= " AND ProdutosConsulta.Situacao not in(1) " & vbCrLf
        End If

        Sql &= " " & vbCrLf & _
        "GROUP BY Consulta.Empresa, Consulta.EndEmpresa, " & vbCrLf & _
        "       Consulta.Deposito, Consulta.EndDeposito, " & vbCrLf & _
        "       Depositos.Nome, Depositos.Cidade, " & vbCrLf & _
        "       Depositos.Estado, Depositos.Reduzido, " & vbCrLf & _
        "       ProdutosConsulta.Grupo, GruposDeEstoques.Descricao, " & vbCrLf & _
        "       Consulta.Produto, ProdutosConsulta.Nome, " & vbCrLf & _
        "       Consulta.Movimento " & vbCrLf & _
        " " & vbCrLf & _
        " " & vbCrLf & _
        "select Deposito, EndDeposito, NomeDoDeposito, CidadeDoDeposito, EstadoDoDeposito, ReduzidoDoDeposito, " & vbCrLf & _
        "		Grupo, NomeDoGrupo, Produto, NomeDoProduto, Movimento, Entradas, Saidas, Producao, Consumo " & vbCrLf & _
        "from #Produto" & vbCrLf & _
        "ORDER BY Deposito, EndDeposito, Produto, Movimento" & vbCrLf & _
        " " & vbCrLf & _
        " " & vbCrLf & _
        "select Produto, NomeDoProduto, sum(Entradas) as Entradas, sum(Producao) As Producao, sum(Consumo) as Consumo, sum(Saidas) as Saidas " & vbCrLf & _
        "from #Produto " & vbCrLf & _
        "group by Produto, NomeDoProduto " & vbCrLf & _
        "order by NomeDoProduto"

        Return Sql
    End Function

    Private Sub BuscaUnidadeNegocio()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
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
        ucSelecaoProduto.Limpar()

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
        BuscaEmpresa()
    End Sub


    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            Relatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            Relatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Relatorio(ByVal pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("ComposicaoDoEstoque", "RELATORIO") Then
                If Validar() Then

                    If rdFisico.Checked Then
                        Sql = RelatorioFisico()
                    Else
                        Sql = RelatorioFiscal()
                    End If

                    ds = Banco.ConsultaDataSet(Sql, "ComposicaoDoEstoque")

                    If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Sem informações para os parâmetros selecionados.", eTitulo.Info)
                        Exit Sub
                    End If

                    If pdf Then
                        Dim parameters = New Dictionary(Of String, Object)()
                        parameters.Add("Parametros", getParam())

                        Funcoes.BindReport(Me.Page, ds, "Cr_ComposicaoDoEstoque", eExportType.PDF, parameters)
                    Else
                        EmitirRelatorioExcel()
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorioExcel()
        Dim objEmpresa As New [Lib].Negocio.Cliente(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1))

        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
            Using package As New ExcelPackage(arquivo)

                'criando planilha títulos

                '******************************************
                '*** 1 - 'Posição Diária
                '******************************************
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Posição Diária")

                'criando linha com o cabeçalho da planilha
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1

                'criando linha que informa o nome da empresa e o cnpj
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa a cidade e o estado da empresa
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o título do relatório
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO DIÁRIA DE ESTOQUE")
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o período selecionado na página
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha com o cabeçalho da planilha
                For Each col As DataColumn In ds.Tables(0).Columns
                    If col.ColumnName = "Produto" Then
                        worksheet.Cells(rowIndex, 1).Value = "PRODUTO"
                        worksheet.Cells(rowIndex, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, 1).Style.Font.Size = 12
                    ElseIf col.ColumnName = "NomeDoProduto" Then
                        worksheet.Cells(rowIndex, 2).Value = "NOME"
                        worksheet.Cells(rowIndex, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, 2).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Movimento" Then
                        worksheet.Cells(rowIndex, 3).Value = "MOVIMENTO"
                        worksheet.Cells(rowIndex, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, 3).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Entradas" Then
                        worksheet.Cells(rowIndex, 4).Value = "ENTRADA"
                        worksheet.Cells(rowIndex, 4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, 4).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Producao" Then
                        worksheet.Cells(rowIndex, 5).Value = "PRODUÇÃO"
                        worksheet.Cells(rowIndex, 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, 5).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Consumo" Then
                        worksheet.Cells(rowIndex, 6).Value = "CONSUMO"
                        worksheet.Cells(rowIndex, 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, 6).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Saidas" Then
                        worksheet.Cells(rowIndex, 7).Value = "SAÍDA"
                        worksheet.Cells(rowIndex, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, 7).Style.Font.Size = 12
                    End If
                Next

                'criando auto filtro na planilha
                worksheet.Cells("A5:C" & rowIndex).AutoFilter = True

                'aplicando formatação nas células do cabeçalho
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 8)
                    range.Style.Font.Bold = True
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    range.Style.Font.Color.SetColor(Color.White)
                End Using
                rowIndex += 1

                Dim descricao As String = String.Empty
                Dim primeiraVez As Boolean = True
                Dim rowINI As Integer

                For Each row As DataRow In ds.Tables(0).Rows
                    columnIndex = 1

                    If Not descricao = row.Item(0) Then
                        If primeiraVez Then
                            primeiraVez = False
                        Else
                            'criando colunas de totalizadores
                            worksheet.Cells(String.Format("C{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")
                            worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Bold = True

                            worksheet.Cells(String.Format("D{0}", rowIndex)).Formula = String.Format("=SUM(D{0}:D{1})", rowINI, rowIndex - 1)
                            worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                            worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Bold = True

                            worksheet.Cells(String.Format("E{0}", rowIndex)).Formula = String.Format("=SUM(E{0}:E{1})", rowINI, rowIndex - 1)
                            worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                            worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Bold = True

                            worksheet.Cells(String.Format("F{0}", rowIndex)).Formula = String.Format("=SUM(F{0}:F{1})", rowINI, rowIndex - 1)
                            worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                            worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True

                            worksheet.Cells(String.Format("G{0}", rowIndex)).Formula = String.Format("=SUM(G{0}:G{1})", rowINI, rowIndex - 1)
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True

                            rowIndex += 2
                        End If

                        descricao = row.Item(0)

                        worksheet.Cells(rowIndex, columnIndex).Value = Funcoes.FormatarCpfCnpj(row.Item(0)) & "-" & row.Item(1) & " - " & row.Item(2)
                        worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 10
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True

                        rowIndex += 1

                        rowINI = rowIndex
                    End If

                    'Produto
                    worksheet.Cells(rowIndex, 1).Value = row.Item(8)
                    'Nome
                    worksheet.Cells(rowIndex, 2).Value = row.Item(9)
                    'Movimento
                    worksheet.Cells(rowIndex, 3).Value = row.Item(10)
                    'Entrada
                    worksheet.Cells(rowIndex, 4).Value = row.Item(11)
                    'Producao
                    worksheet.Cells(rowIndex, 5).Value = row.Item(13)
                    'Consumo
                    worksheet.Cells(rowIndex, 6).Value = row.Item(14)
                    'Saida
                    worksheet.Cells(rowIndex, 7).Value = row.Item(12)

                    'formatando células datas
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"

                    rowIndex += 1
                Next

                'criando colunas de totalizadores
                worksheet.Cells(String.Format("C{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")
                worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("D{0}", rowIndex)).Formula = String.Format("=SUM(D{0}:D{1})", rowINI, rowIndex - 1)
                worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("E{0}", rowIndex)).Formula = String.Format("=SUM(E{0}:E{1})", rowINI, rowIndex - 1)
                worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("F{0}", rowIndex)).Formula = String.Format("=SUM(F{0}:F{1})", rowINI, rowIndex - 1)
                worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("G{0}", rowIndex)).Formula = String.Format("=SUM(G{0}:G{1})", rowINI, rowIndex - 1)
                worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True

                'setando autofit nas células da planilha
                worksheet.Cells.AutoFitColumns(0)

                'congelando quinta linha (cabeçalho)
                worksheet.View.FreezePanes(6, 1)

                '******************************************
                '*** 2 - 'Resumo
                '******************************************
                worksheet = package.Workbook.Worksheets.Add("Resumo")
                worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                'criando linha com o cabeçalho da planilha
                rowIndex = 1
                columnIndex = 1

                'criando linha que informa o nome da empresa e o cnpj
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa a cidade e o estado da empresa
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o título do relatório
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RESUMO POR PRODUTO")
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o período selecionado na página
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:B{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha com o cabeçalho da planilha
                For Each col As DataColumn In ds.Tables(1).Columns
                    If col.ColumnName = "Produto" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "PRODUTO"
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                    ElseIf col.ColumnName = "NomeDoProduto" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "NOME"
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Entradas" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "ENTRADA"
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Producao" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "PRODUÇÃO"
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Consumo" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "CONSUMO"
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                    ElseIf col.ColumnName = "Saidas" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "SAÍDA"
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                    End If

                    columnIndex += 1
                Next

                'criando auto filtro na planilha
                worksheet.Cells("A5:B" & rowIndex).AutoFilter = True

                'aplicando formatação nas células do cabeçalho
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count)
                    range.Style.Font.Bold = True
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    range.Style.Font.Color.SetColor(Color.White)
                End Using
                rowIndex += 1


                'criando conteúdo da planilha com os dados do dataset
                For Each row As DataRow In ds.Tables(1).Rows
                    columnIndex = 1

                    For Each col As DataColumn In ds.Tables(1).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                        columnIndex += 1
                    Next

                    'formatando células numéricas
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"

                    'aplicando formatação nas células do conteúdo
                    If rowIndex Mod 2 = 0 Then
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                        End Using
                    End If

                    rowIndex += 1
                Next

                rowIndex += 1

                'criando colunas de totalizadores
                worksheet.Cells(String.Format("B{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")
                worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("C{0}", rowIndex)).Formula = String.Format("=SUM(C6:C{0})", rowIndex - 1)
                worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("D{0}", rowIndex)).Formula = String.Format("=SUM(D6:D{0})", rowIndex - 1)
                worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("E{0}", rowIndex)).Formula = String.Format("=SUM(E6:E{0})", rowIndex - 1)
                worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Bold = True

                worksheet.Cells(String.Format("F{0}", rowIndex)).Formula = String.Format("=SUM(F6:F{0})", rowIndex - 1)
                worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True

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
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ComposicaoDoEstoque")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class