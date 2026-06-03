Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PedidosPendentesDeCompras
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Mensagem As String
    Dim Cliente() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PedidosPendentesDeCompras", "ACESSAR") Then
                    CargaUnidade()
                    VerificaUnidade()
                    CargaMes()
                    CargaAno()
                    ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                    HID.Value = Guid.NewGuid().ToString
                    gridCli.Visible = True
                    gridPrd.Visible = False
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaMes()
        ddlMes.Items.Add(New ListItem("Janeiro", 1))
        ddlMes.Items.Add(New ListItem("Fevereiro", 2))
        ddlMes.Items.Add(New ListItem("Março", 3))
        ddlMes.Items.Add(New ListItem("Abril", 4))
        ddlMes.Items.Add(New ListItem("Maio", 5))
        ddlMes.Items.Add(New ListItem("Junho", 6))
        ddlMes.Items.Add(New ListItem("Julho", 7))
        ddlMes.Items.Add(New ListItem("Agosto", 8))
        ddlMes.Items.Add(New ListItem("Setembro", 9))
        ddlMes.Items.Add(New ListItem("Outubro", 10))
        ddlMes.Items.Add(New ListItem("Novembro", 11))
        ddlMes.Items.Add(New ListItem("Dezembro", 12))
    End Sub

    Private Sub CargaAno()
        ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)

        ddlAno.SelectedValue = Format(Today, "yyyy")
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " &
              "FROM Clientes C " &
              "INNER JOIN ClientesXTipos CT " &
              "ON C.Cliente_Id = CT.Cliente_Id " &
              "WHERE CT.Tipo_Id = 050 " &
              "ORDER BY Nome"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade,"
        Sql &= "        isnull(AcessoEmpresa, '') as AcessoEmpresa,"
        Sql &= "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa"
        Sql &= " from Usuarios"
        Sql &= " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
        Sql &= " FROM   GruposXEmpresas INNER JOIN"
        Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
        Sql &= " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0

    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""

        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()

        ucSelecaoProduto.Limpar()

        GridProduto.DataSource = Nothing
        GridProduto.DataBind()

        GridCliente.DataSource = Nothing
        GridCliente.DataBind()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function Validar()
        Mensagem = ""

        If DdlUnidade.Text = "" Then
            Mensagem = "Unidade de negócio é obrigatório."
            Return Mensagem
        End If
        If DdlEmpresa.Text = "" Then
            Mensagem = "Empresa é obrigatório."
            Return Mensagem
        End If

        Return Mensagem
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Session("objClienteCxN" & HID.Value) IsNot Nothing Then
                Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteCxN" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteCxN" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCxN" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            Validar()
            If Mensagem = "" Then

                Dim DataInicial As String = ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-01"

                GridProduto.DataSource = Nothing
                GridProduto.DataBind()

                GridCliente.DataSource = Nothing
                GridCliente.DataBind()

                Cliente = txtCodigoCliente.Value.Split("-")

                If RadCliente.Checked = True Then
                    Sql = "SELECT Consulta.Pedido_Id AS Pedido," & vbCrLf & _
                          "       PedidosConsulta.DataPedido AS Data," & vbCrLf & _
                          "       Clientes.Nome AS Cliente, Clientes.Bairro, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
                          "       SUM(Consulta.QuantidadePedido - Consulta.QuantidadeEntregue) AS Quantidade," & vbCrLf & _
                          "       SUM(Consulta.ValorPedido - Consulta.ValorNota) AS Valor" & vbCrLf & _
                          "  FROM (" & vbCrLf & _
                          "        SELECT Pedidos.Pedido_Id," & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN isnull(NotasFiscaisXItens.EntradaSaida_Id, 'E') = 'E'" & vbCrLf & _
                          "                       THEN pxi.Quantidade" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
                          "                   END) AS QuantidadePedido, " & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN NotasFiscaisXItens.EntradaSaida_Id = 'S'" & vbCrLf & _
                          "                       THEN ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0)" & vbCrLf & _
                          "                   END) AS QuantidadeEntregue, " & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN isnull(NotasFiscaisXItens.EntradaSaida_Id, 'E') = 'E'" & vbCrLf & _
                          "                       THEN pxi.TotalOficial" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.Valor, 0) * - 1" & vbCrLf & _
                          "                   END) AS ValorPedido, " & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN NotasFiscaisXItens.EntradaSaida_Id = 'S'" & vbCrLf & _
                          "                       THEN ISNULL(NotasFiscaisXItens.Valor, 0) * - 1" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.Valor, 0)" & vbCrLf & _
                          "                   END) AS ValorNota" & vbCrLf & _
                          "          FROM Pedidos" & vbCrLf & _
                          "         INNER JOIN VW_PedidosXItensXFixacoes pxi" & vbCrLf & _
                          "            ON Pedidos.Empresa_Id    = pxi.Empresa_Id" & vbCrLf & _
                          "           AND Pedidos.EndEmpresa_Id = pxi.EndEmpresa_Id" & vbCrLf & _
                          "           AND Pedidos.Pedido_Id     = pxi.Pedido_Id" & vbCrLf & _
                          "         INNER JOIN SubOperacoes" & vbCrLf & _
                          "            ON Pedidos.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
                          "           AND Pedidos.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                          "          LEFT JOIN NotasFiscaisXItens" & vbCrLf & _
                          "            ON pxi.Empresa_Id    = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                          "           AND pxi.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                          "           AND pxi.Pedido_Id     = NotasFiscaisXItens.Pedido" & vbCrLf & _
                          "           AND pxi.Produto_Id    = NotasFiscaisXItens.Produto_Id" & vbCrLf & _
                          "         WHERE Pedidos.DataPedido BETWEEN '2010-01-01' AND '2015-12-31'" & vbCrLf & _
                          "           AND SubOperacoes.EntradaSaida = 'E'" & vbCrLf & _
                          "           AND (Pedidos.Situacao = 1)" & vbCrLf

                    If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                        Sql &= "         And Pedidos.Cliente = '" & Cliente(0) & "' And Pedidos.EndCliente = " & Cliente(1) & vbCrLf
                    End If

                    If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND Pedidos.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf

                    If cmbSubOperacao.SelectedIndex > 0 Then
                        Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
                        Sql &= "AND Pedidos.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
                    End If

                    Sql &= "         GROUP BY   Pedidos.Pedido_Id" & vbCrLf & _
                          "       ) AS Consulta " & vbCrLf & _
                          " INNER JOIN Pedidos AS PedidosConsulta" & vbCrLf & _
                          "    ON Consulta.Pedido_Id = PedidosConsulta.Pedido_Id " & vbCrLf & _
                          "  LEFT JOIN Clientes" & vbCrLf & _
                          "    ON PedidosConsulta.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                          "   AND PedidosConsulta.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                          " WHERE Consulta.QuantidadePedido <> Consulta.QuantidadeEntregue" & vbCrLf & _
                          " GROUP BY Consulta.Pedido_Id, PedidosConsulta.DataPedido, PedidosConsulta.PedidoEfetivo, Clientes.Nome, Clientes.Cidade, Clientes.Bairro, Clientes.Estado" & vbCrLf & _
                          " ORDER BY   Cliente, Data, Pedido" & vbCrLf
                    '#NGSVerificar  Data Fixa KD a Selecao dos Produtos?
                    GridCliente.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
                    GridCliente.DataBind()
                Else

                    Sql = "SELECT Consulta.Pedido_Id AS Pedido, PedidosConsulta.DataPedido AS Data, Clientes.Nome AS Cliente, Produto_Id as Produto, Unidade, NomeDoProduto," & vbCrLf & _
                          "       Consulta.QuantidadePedido - Consulta.QuantidadeEntregue AS Quantidade," & vbCrLf & _
                          "       Unitario," & vbCrLf & _
                          "       Consulta.ValorPedido - Consulta.ValorNota AS Valor" & vbCrLf & _
                          "  FROM (" & vbCrLf & _
                          "        SELECT Pedidos.Pedido_Id," & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN isnull(NotasFiscaisXItens.EntradaSaida_Id, 'E') = 'E'" & vbCrLf & _
                          "                       THEN pxi.Quantidade" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
                          "                   END) AS QuantidadePedido, " & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN NotasFiscaisXItens.EntradaSaida_Id = 'S'" & vbCrLf & _
                          "                       THEN ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.QuantidadeFiscal, 0)" & vbCrLf & _
                          "                   END) AS QuantidadeEntregue, " & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN isnull(NotasFiscaisXItens.EntradaSaida_Id, 'E') = 'E' " & vbCrLf & _
                          "                       THEN pxi.TotalOficial" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.Valor, 0) * - 1" & vbCrLf & _
                          "                   END) AS ValorPedido," & vbCrLf & _
                          "               SUM(CASE" & vbCrLf & _
                          "                     WHEN NotasFiscaisXItens.EntradaSaida_Id = 'S'" & vbCrLf & _
                          "                       THEN ISNULL(NotasFiscaisXItens.Valor, 0) * - 1" & vbCrLf & _
                          "                       ELSE ISNULL(NotasFiscaisXItens.Valor, 0)" & vbCrLf & _
                          "                   END) AS ValorNota," & vbCrLf & _
                          "               Produtos.Produto_Id," & vbCrLf & _
                          "               Produtos.Unidade," & vbCrLf & _
                          "               Produtos.Nome AS NomeDoProduto, " & vbCrLf & _
                          "               pxi.UnitarioOficial as Unitario" & vbCrLf & _
                          "          FROM Pedidos" & vbCrLf & _
                          "         INNER JOIN VW_PedidosXItensXFixacoes pxi" & vbCrLf & _
                          "            ON Pedidos.Empresa_Id    = pxi.Empresa_Id" & vbCrLf & _
                          "           AND Pedidos.EndEmpresa_Id = pxi.EndEmpresa_Id  " & vbCrLf & _
                          "           AND Pedidos.Pedido_Id     = pxi.Pedido_Id " & vbCrLf & _
                          "         INNER JOIN SubOperacoes" & vbCrLf & _
                          "            ON Pedidos.Operacao = SubOperacoes.Operacao_Id" & vbCrLf & _
                          "           AND Pedidos.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                          "         INNER JOIN Produtos" & vbCrLf & _
                          "            ON pxi.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                          "          LEFT JOIN NotasFiscaisXItens" & vbCrLf & _
                          "            ON pxi.Empresa_Id    = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                          "           AND pxi.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                          "           AND pxi.Pedido_Id     = NotasFiscaisXItens.Pedido" & vbCrLf & _
                          "           AND pxi.Produto_Id    = NotasFiscaisXItens.Produto_Id" & vbCrLf & _
                          "         WHERE Pedidos.DataPedido BETWEEN '2010-01-01' AND '2015-12-31'" & vbCrLf & _
                          "           AND SubOperacoes.EntradaSaida = 'E'" & vbCrLf & _
                          "           AND Pedidos.Situacao          = 1" & vbCrLf
                    '#NGSVerificar Data Fixa
                    If txtCliente.Text <> "" Then
                        Sql &= "  And (Pedidos.Cliente = '" & Cliente(0) & "') And Pedidos.EndCliente = " & Cliente(1)
                    End If

                    If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND Pedidos.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf

                    If cmbSubOperacao.SelectedIndex > 0 Then
                        Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
                        Sql &= "AND Pedidos.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
                    End If

                    If ucSelecaoProduto.TemSelecionado Then
                        Dim RetornoProdutos As ArrayList
                        RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "pxi.Produto_Id", "")
                        Sql &= " AND " & RetornoProdutos(0)
                    End If

                    Sql &= "         GROUP BY Pedidos.Pedido_Id, Produtos.Produto_Id, Produtos.Unidade, Produtos.Nome, pxi.UnitarioOficial" & vbCrLf & _
                           "       ) AS Consulta" & vbCrLf & _
                           " INNER JOIN Pedidos AS PedidosConsulta" & vbCrLf & _
                           "    ON Consulta.Pedido_Id = PedidosConsulta.Pedido_Id" & vbCrLf & _
                           "  LEFT JOIN Clientes" & vbCrLf & _
                           "    ON PedidosConsulta.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                           "   AND PedidosConsulta.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                           " WHERE Consulta.QuantidadePedido <> Consulta.QuantidadeEntregue" & vbCrLf & _
                           " ORDER BY Cliente, Data, Pedido" & vbCrLf
                    GridProduto.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
                    GridProduto.DataBind()
                End If
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            Funcoes.Ajuda(Me.Page, "PedidosPendentesDeCompras")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub RadCliente_CheckedChanged(sender As Object, e As EventArgs) Handles RadCliente.CheckedChanged
        gridCli.Visible = True
        gridPrd.Visible = False

        GridProduto.DataSource = Nothing
        GridProduto.DataBind()
    End Sub

    Protected Sub RadProduto_CheckedChanged(sender As Object, e As EventArgs) Handles RadProduto.CheckedChanged
        gridCli.Visible = False
        gridPrd.Visible = True

        GridCliente.DataSource = Nothing
        GridCliente.DataBind()
    End Sub
End Class