Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ComprasVendasAFixar
    Inherits BasePage

    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ComprasVendasAFixar", "ACESSAR") Then
                TxtData.Text = Format(Today, "dd/MM/yyyy")
                CargaUnidade()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa)
        ucSelecaoProduto.Limpar()
        TxtData.Text = Format(Today, "dd/MM/yyyy")
        RdCompras.Checked = False
        RdVendas.Checked = False
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function Validar()
        If ddlUnidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf DdlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Parametros:" & vbCrLf & "Unidade: " & ddlUnidade.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(TxtData.Text) Then
            param &= "Até a Data: " & TxtData.Text & " "
        End If
        param &= IIf(RdCompras.Checked, "Selecionado: Compras", "Selecionado: Vendas")

        Return param
    End Function

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function Relatorio() As DataSet
        Dim Ds As New DataSet
        Dim strOR As String = ""


        Sql = "SELECT sdAf.Empresa_Id AS Empresa, sdAf.EndEmpresa_Id AS EndEmpresa, sdAf.Cliente_Id AS Cliente, sdAf.EndCliente_Id AS EndCliente, " & vbCrLf & _
                "		sdAf.Produto_Id AS Produto, sdAf.Pedido, sdAf.QuantidadeFiscal AS Afixar, 0 AS Fixado, 0 AS Saldo " & vbCrLf & _
                "into #Temp " & vbCrLf & _
                "From (SELECT nXi.Empresa_Id, nXi.EndEmpresa_Id, nXi.Cliente_Id, " & vbCrLf & _
                "						nXi.EndCliente_Id, nXi.Pedido, nXi.Produto_Id, " & vbCrLf & _
                "						sum(Case " & vbCrLf & _
                "							when nXi.EntradaSaida_Id = 'S' " & vbCrLf & _
                "								then nXi.QuantidadeFiscal * -1 " & vbCrLf & _
                "								else nXi.QuantidadeFiscal " & vbCrLf & _
                "						end) AS QuantidadeFiscal " & vbCrLf & _
                "				From NotasFiscaisXitens nXi " & vbCrLf & _
                "					INNER JOIN NotasFiscais n " & vbCrLf & _
                "							 on n.empresa_id      = nXi.empresa_id " & vbCrLf & _
                "							and n.endempresa_id   = nXi.endempresa_id " & vbCrLf & _
                "							and n.Cliente_Id      = nXi.Cliente_Id " & vbCrLf & _
                "							and n.EndCliente_Id   = nXi.EndCliente_Id " & vbCrLf & _
                "							and n.EntradaSaida_Id = nXi.EntradaSaida_Id " & vbCrLf & _
                "							and n.Serie_Id        = nXi.Serie_Id " & vbCrLf & _
                "							and n.Nota_Id         = nXi.Nota_Id " & vbCrLf & _
                "					INNER JOIN SubOperacoes SO " & vbCrLf & _
                "							 ON SO.Operacao_Id     = nXi.Operacao " & vbCrLf & _
                "							AND SO.SubOperacoes_Id = nXi.SubOperacao " & vbCrLf & _
                "							AND (SO.Classe = '" & eClassesOperacoes.AFIXAR.ToString & "') " & vbCrLf & _
                "				WHERE (n.Movimento <= '" & TxtData.Text.ToSqlDate() & "') " & vbCrLf

        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList
            RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "nXi.Produto_Id", "")
            Sql &= " AND " & RetornoProdutos(0)
        End If

        Sql &= "				  AND (nXi.EntradaSaida_Id = '" & IIf(RdCompras.Checked, "E", "S") & "') " & vbCrLf & _
               "				  AND (SO.Classe = '" & eClassesOperacoes.AFIXAR.ToString & "') " & vbCrLf & _
               "				GROUP BY nXi.Empresa_Id, nXi.EndEmpresa_Id, nXi.Cliente_Id, nXi.EndCliente_Id, " & vbCrLf & _
               "						 nXi.Pedido, nXi.Produto_Id) AS sdAf " & vbCrLf & _
               " update #Temp set" & vbCrLf & _
               "	 Fixado = isnull(pFx.PesoFixado,0), Saldo = (Afixar - isnull(pFx.PesoFixado,0)) " & vbCrLf & _
               "   from #Temp " & vbCrLf & _
               "   left join(SELECT Empresa_id, EndEmpresa_Id, Pedido_Id, Produto_Id, sum(Quantidade) AS PesoFixado " & vbCrLf & _
               "   			   FROM VW_PedidosXItensXFixacoes " & vbCrLf & _
               "   			  group by Empresa_id, EndEmpresa_Id, Pedido_Id, Produto_Id" & vbCrLf & _
               "            ) AS pFx " & vbCrLf & _
               "   	 ON pFx.Empresa_Id    = Empresa " & vbCrLf & _
               "   	AND pFx.EndEmpresa_id = EndEmpresa " & vbCrLf & _
               "   	AND pFx.Pedido_Id     = Pedido " & vbCrLf & _
               "   	AND pFx.Produto_Id    = Produto_Id;" & vbCrLf & _
               "Select Empresa, EndEmpresa, Cliente, EndCliente, Produto, sum(Afixar) AS AFixar, sum(Fixado) AS Fixado, sum(Saldo) AS SaldoFinal " & vbCrLf & _
               "  into #Afixar " & vbCrLf & _
               "  from #Temp " & vbCrLf & _
               " group by Empresa, EndEmpresa, Cliente, EndCliente, Produto;" & vbCrLf & _
               "Select Empresa, EndEmpresa, Emp.Nome AS NomeEmpresa, Emp.Cidade AS CidadeEmpresa, Emp.Estado AS EstadoEmpresa, " & vbCrLf & _
               "	   Cliente, EndCliente, Cli.Reduzido, Cli.Nome AS NomeCliente, Cli.Cidade AS CidadeCliente, Cli.Estado AS EstadoCliente, " & vbCrLf & _
               "	   Produto, Prd.Nome AS NomeProduto, AFixar, Fixado, SaldoFinal " & vbCrLf & _
               "  from #Afixar " & vbCrLf & _
               " INNER JOIN Produtos AS Prd " & vbCrLf & _
               "    ON Prd.Produto_Id = Produto " & vbCrLf & _
               " INNER JOIN Clientes AS Emp " & vbCrLf & _
               "    ON Emp.Cliente_Id  = Empresa " & vbCrLf & _
               "   AND Emp.Endereco_Id = EndEmpresa " & vbCrLf & _
               " INNER JOIN Clientes Cli " & vbCrLf & _
               "    ON Cli.Cliente_Id  = Cliente " & vbCrLf & _
               "   AND Cli.Endereco_Id = EndCliente " & vbCrLf

        If ChkZerados.Checked = False Then
            Sql &= "WHERE SaldoFinal > 0 " & vbCrLf
        Else
            Sql &= "WHERE Afixar > 0 " & vbCrLf
        End If

        Sql &= "ORDER BY Produto, Empresa, EndEmpresa, Cli.Nome" & vbCrLf

        Ds = Banco.ConsultaDataSet(Sql, "ComprasVendasAFixar")

        Return Ds

    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ComprasVendasAFixar", "RELATORIO") Then
                If Validar() Then
                    Dim ds As New DataSet
                    ds = Relatorio()

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Titulo", IIf(RdCompras.Checked, "Compras à Fixar.", "Vendas à Fixar."))
                    parameters.Add("Parametros", getParam())

                    Funcoes.BindReport(Me.Page, ds, "Cr_ComprasVendasAFixar", eExportType.PDF, parameters)
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
            Funcoes.Ajuda(Me.Page, "ComprasVendasAFixar")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class