Imports NGS.Lib.Negocio
Imports System.Threading
Imports System.IO

Public Class AjustarPedido
    Inherits BasePage
    Dim objPedido As Pedido

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If IsConnect AndAlso Not IsPostBack Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Comercial.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("AjustarPedido", "ACESSAR") Then
                    CargaUnidadeDeNegocio()
                    ddl.Carregar(ddlOperacao, CarregarDDL.Tabela.Operacao)
                    CargaSubOperacao()

                    limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "Comercial.aspx", eTitulo.Info)
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)

        If Not Session("objClientePXI" & HID.Value) Is Nothing Then
            Dim cli As Cliente = Session("objClientePXI" & HID.Value)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClientePXI" & HID.Value)

            recuperaPedido()

            If Not objPedido Is Nothing AndAlso objPedido.IUD = "U" Then

                objPedido.CodigoCliente = txtCodigoCliente.Value.Split("-")(0)
                objPedido.EnderecoCliente = txtCodigoCliente.Value.Split("-")(1)

                salvarPedido()

            Else
                BuscarPedido()
            End If

        ElseIf Not Session("objClientePedido" & HID.Value) Is Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objClientePedido" & HID.Value), [Lib].Negocio.Pedido)

            txtPedido.Text = p.Codigo

            Session.Remove("objClientePedido" & HID.Value)

            atualizarGrid()
        End If

    End Sub

    Private Sub BuscarPedido()
        Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
        Dim parameters As New Dictionary(Of String, Object)
        parameters("unidade") = ""
        parameters("empresa") = strEmpresa(0)
        parameters("enderecoEmpresa") = strEmpresa(1)
        parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
        parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
        parameters("situacao") = 1
        Popup.ConsultaDePedidos(Me.Page, "objClientePedido" & HID.Value)
        ucConsultaPedidos.BindGridView(parameters)
    End Sub

    Private Sub limpar()
        Session.Remove("objPedido" & HID.Value)
        Session.Remove("objPedVencimentos" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        lnkConfirmar.Parent.Visible = False

        txtCliente.Enabled = True
        txtPedido.Enabled = True
        ddlOperacao.Enabled = False
        ddlSubOperacao.Enabled = False
        ddlOperacao.SelectedValue = String.Empty
        ddlSubOperacao.SelectedValue = String.Empty

        txtCliente.Text = String.Empty
        txtPedido.Text = String.Empty
        txtAnterior.Text = String.Empty
        txtAtual.Text = String.Empty

        grdItens.DataSource = Nothing
        grdItens.DataBind()
        gridEncargosGerais.DataSource = Nothing
        gridEncargosGerais.DataBind()
        grdTitulos.DataSource = Nothing
        grdTitulos.DataBind()

        ucPedidoxEncargo.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaSubOperacao()
        If String.IsNullOrWhiteSpace(ddlOperacao.SelectedValue) Then
            ddlSubOperacao.Items.Clear()
        Else
            ddl.Carregar(ddlSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & ddlOperacao.SelectedValue & " ", True)
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub recuperaPedido()
        objPedido = Session("objPedido" & HID.Value)
    End Sub

    Private Sub salvarPedido()
        Session("objPedido" & HID.Value) = objPedido
    End Sub

    Private Sub atualizarGrid()

        objPedido = New Pedido(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtPedido.Text)

        If objPedido Is Nothing OrElse objPedido.Itens.Count = 0 Then
            MsgBox(Me.Page, "Pedido não encontrado!", eTitulo.Info)
            limpar()
            Exit Sub
        End If

        If txtCliente.Text.Length = 0 Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objPedido.Cliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
        End If

        If objPedido.PedidoBloqueado Then
            MsgBox(Me.Page, "Pedido LIBERADO não pode ser ajustado por aqui!", eTitulo.Info)
            limpar()
            Exit Sub
        End If

        If (Left(objPedido.CodigoEmpresa, 8) = "24450490" OrElse Left(objPedido.CodigoEmpresa, 8) = "44979506") AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso Funcoes.VerificaPermissao("AjustarPedido", "LIBERAR") Then
            'LIBERA PARA AJUSTAR COMPLEMENTO EM DÓLAR - FURLAN - 04/11/2024
            'LIBERADO PARA VERDE - 02/05/2025 - FURLAN
        Else
            If objPedido.TemFinanceiro(1) OrElse objPedido.TemFinanceiro(2) Then
                MsgBox(Me.Page, "Pedido com faturamento não pode ser modificado, apenas estornar ou complementar.", eTitulo.Info)
                limpar()
                Exit Sub
            ElseIf objPedido.TemFaturamento Then
                MsgBox(Me.Page, "Pedido com nota fiscal não pode ser modificado, apenas estornar ou complementar.", eTitulo.Info)
                limpar()
                Exit Sub
            End If

            For Each item In objPedido.Itens
                If item.Lancamentos.Count > 1 Then
                    MsgBox(Me.Page, "Produto com estorno ou complemento, não pode ser alterado.")
                    limpar()
                    Exit Sub
                End If
            Next
        End If

        If objPedido.TemNFG Then
            MsgBox(Me.Page, "Pedido de Notas Gerais não pode ser ajustado por aqui!", eTitulo.Info)
            limpar()
            Exit Sub
        End If

        If objPedido.Vencimentos IsNot Nothing AndAlso objPedido.Vencimentos.Count > 0 Then
            Dim objVencimentos As New Hashtable
            Dim i As Integer
            For i = 0 To objPedido.Vencimentos.OrderBy(Function(s) s.Provisao).Count - 1
                objVencimentos.Add(i, objPedido.Vencimentos(i).Codigo)
            Next
            Session("objPedVencimentos" & HID.Value) = objVencimentos
        End If

        objPedido.IUD = "U"

        ddlUnidadeDeNegocio.Enabled = False
        ddlEmpresa.Enabled = False
        txtCliente.Enabled = False
        txtPedido.Enabled = False

        salvarPedido()

        grdItens.DataSource = objPedido.Itens.ToArray
        grdItens.DataBind()

        If (Left(objPedido.CodigoEmpresa, 8) = "24450490" OrElse Left(objPedido.CodigoEmpresa, 8) = "44979506") AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso Funcoes.VerificaPermissao("AjustarPedido", "LIBERAR") Then
            grdItens.Columns(7).Visible = True
            grdItens.Columns(8).Visible = True
        Else
            grdItens.Columns(7).Visible = False
            grdItens.Columns(8).Visible = False
        End If

        AtualizarGridEncargos(False)

        If objPedido.Vencimentos IsNot Nothing AndAlso objPedido.Vencimentos.Count > 0 Then
            grdTitulos.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
            grdTitulos.DataBind()
        End If

        txtAnterior.Text = totalPedido()
        txtAtual.Text = totalPedido()

        ddlOperacao.SelectedValue = objPedido.CodigoOperacao
        CargaSubOperacao()
        ddlSubOperacao.SelectedValue = objPedido.CodigoOperacao & "-" & objPedido.CodigoSubOperacao
        ddlOperacao.Enabled = True
        ddlSubOperacao.Enabled = True
    End Sub

    Protected Sub ddlOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOperacao.SelectedIndexChanged
        Try
            CargaSubOperacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlSubOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSubOperacao.SelectedIndexChanged
        Try
            If ddlSubOperacao.SelectedValue.ToString.Split("-")(0) > 0 Then
                recuperaPedido()

                objPedido.CodigoOperacao = ddlSubOperacao.SelectedValue.ToString.Split("-")(0)
                objPedido.CodigoSubOperacao = ddlSubOperacao.SelectedValue.ToString.Split("-")(1)

                For Each Item In objPedido.Itens
                    Dim dataAnteriorItem As Date = Item.Lancamentos.LancamentoNormal.Movimento
                    Item.IUD = "U"
                    Item.Lancamentos.LancamentoNormal.Movimento = objPedido.DataPedido
                    Item.Encargos = Nothing
                    Item.CodigoOperacaoXEstado = 0
                    Item.Encargos.CriaListar()
                    If Not Item.Encargos.Count() > 0 Then
                        limpar()
                        MsgBox(Me.Page, "Não foram encontrado encargos para operação.", eTitulo.Info)
                        Exit Sub
                    End If
                    Item.Encargos.AtualizaLiquido()
                    Item.Lancamentos.LancamentoNormal.Movimento = dataAnteriorItem
                Next
            End If

            lnkConfirmar.Parent.Visible = True

            salvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Exit Sub
        End Try
    End Sub

    Private Sub atualizarItem(index As Integer)
        Dim sql As String = ""

        recuperaPedido()

        Dim quantidade As Decimal = CType(grdItens.Rows(index).FindControl("txtQtd"), TextBox).Text()
        Dim unitario As Decimal = CType(grdItens.Rows(index).FindControl("txtUnitario"), TextBox).Text()
        Dim total As Decimal

        For Each item In objPedido.Itens.Where(Function(s) s.CodigoProduto = grdItens.Rows(index).Cells(0).Text)
            item.IUD = "U"
            item.Lancamentos(0).IUD = "U"
            item.Lancamentos(0).QuantidadeFaturamento = quantidade
            item.Lancamentos(0).QuantidadeComercializacao = quantidade
            item.Lancamentos(0).UnitarioOficial = unitario
            item.Lancamentos(0).TotalOficial = Math.Round((quantidade * unitario), 2, MidpointRounding.AwayFromZero)

            total = item.Lancamentos(0).TotalOficial

            If objPedido.IndiceFixado > 0 Then
                item.Lancamentos(0).UnitarioMoeda = Funcoes.ConverteMoeda(item.Lancamentos(0).UnitarioOficial, objPedido.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 10)
            End If

            item.Encargos = Nothing
            item.Encargos.CriaListar()
        Next

        salvarPedido()

        grdItens.Rows(index).Cells(4).Text = total.ToString("N2")

        AtualizarGridEncargos(True)
    End Sub

    Public Sub AtualizarGridEncargos(ByVal liberar As Boolean)

        recuperaPedido()

        gridEncargosGerais.DataSource = From p In objPedido.Itens.SelectMany(Function(s) s.Encargos)
                                        Group By p.CodigoEncargo
                                         Into ValorMoeda = Sum(p.ValorMoeda), ValorOficial = Sum(p.ValorOficial)
                                        Order By IIf(CodigoEncargo = "PRODUTO", 1, IIf(CodigoEncargo = "LIQUIDO", 3, 2))
                                        Select CodigoEncargo, ValorMoeda, ValorOficial

        If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            gridEncargosGerais.Columns(2).Visible = False
        Else
            gridEncargosGerais.Columns(2).Visible = True
        End If
        gridEncargosGerais.DataBind()

        If liberar Then lnkConfirmar.Parent.Visible = True
    End Sub

    Private Sub confirmarAtualizacoes()
        recuperaPedido()

        Dim Sqls As New ArrayList

        objPedido.UsuarioAlteracao = Session("ssNomeUsuario")
        objPedido.DataAlteracao = Now()
        objPedido.SalvarSql(Sqls)
        objPedido.Itens.SalvarSql(Sqls)

        If (objPedido IsNot Nothing AndAlso objPedido.Vencimentos IsNot Nothing AndAlso objPedido.Vencimentos.Count > 0) Then
            objPedido.Vencimentos.CriarParcelamento(False, CType(Session("objPedVencimentos" & HID.Value), Hashtable))
            Sqls.AddRange(objPedido.Vencimentos.GetSQL(eModoAlteracao.Inclusao, UsuarioServidor.NomeServidor.ToUpper(), objPedido.MomentoFinanceiro))
        End If

        Dim verificarPedido As New Pedido(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtPedido.Text)

        'Se houve alteração do cliente
        If String.Format("{0}|{1}", verificarPedido.Cliente.Codigo, verificarPedido.Cliente.CodigoEndereco) <> String.Format("{0}|{1}", objPedido.Cliente.Codigo, objPedido.Cliente.CodigoEndereco) Then

            Dim Sql As String
            Sql = "         UPDATE PedidosXDepositos 
                                SET Deposito_Id             = '" & objPedido.Cliente.Codigo & "',
                                    EndDeposito_Id          = " & objPedido.Cliente.CodigoEndereco & "
                            WHERE Empresa_Id                = '" & verificarPedido.CodigoEmpresa & "'
                                AND EndEmpresa_Id           = " & verificarPedido.EnderecoEmpresa & "
                                AND Pedido_Id               = " & verificarPedido.Codigo & "
                                AND Deposito_Id             = '" & verificarPedido.Cliente.Codigo & "'
                                AND EndDeposito_Id          = " & verificarPedido.Cliente.CodigoEndereco & ""
            Sqls.Add(Sql)


            Sql = " SELECT Registro_id, 'ContasAReceber' AS Tabela
                    FROM ContasAReceber  
                    WHERE 1=1
                        AND Pedido      = " & objPedido.Codigo & " 
                        AND Empresa     = '" & objPedido.CodigoEmpresa & "' 
                        AND EndEmpresa  = " & objPedido.EnderecoEmpresa & " 
                        AND Provisao    = " & eProvisao.Provisao & " 
                        AND Situacao    = 1
                    UNION 
                    SELECT Registro_id, 'ContasAPagar' AS Tabela 
                    FROM ContasAPagar 
                    WHERE 1=1
                        AND Pedido      = " & objPedido.Codigo & " 
                        AND Empresa     = '" & objPedido.CodigoEmpresa & "' 
                        AND EndEmpresa  = " & objPedido.EnderecoEmpresa & " 
                        AND Provisao    = " & eProvisao.Provisao & " 
                        AND Situacao    = 1; "

            Dim ds As DataSet
            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            For Each row As DataRow In ds.Tables(0).Rows

                Sql = " UPDATE " & row("Tabela") & " SET    Cliente                 = '" & objPedido.Cliente.Codigo & "'
                                                            ,EndCliente             = " & objPedido.Cliente.CodigoEndereco & "
                                                            ,Destinatario           = '" & objPedido.Cliente.Codigo & "'
                                                            ,EndDestinatario        = " & objPedido.Cliente.CodigoEndereco & "
                                                            ,NomeDoDestinatario     = '" & objPedido.Cliente.Nome & "'
                            WHERE Registro_Id = " & row("Registro_Id") & "; "

                Sqls.Add(Sql)
            Next

        End If

        If Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, "Pedido ajustado com suscesso!", eTitulo.Sucess)

            objPedido = New Pedido(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtPedido.Text)

            objPedido.ImprimirPedido(Me.Page, True)

            limpar()
        Else
            MsgBox(Me.Page, "Houve um erro ao gravar As alterações, contate o suporte.", eTitulo.Erro)
        End If

    End Sub

    Private Sub confirmarComplemento(index As Integer)
        recuperaPedido()

        Dim Sqls As New ArrayList

        Dim ItemPedido As [Lib].Negocio.PedidoXItem
        ItemPedido = objPedido.Itens(index)
        ItemPedido.IUD = "I"

        Dim Sql As String = "DELETE FROM PedidosXEncargos" & vbCrLf &
                            "Where Empresa_Id    = '" & objPedido.CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf &
                            "  and Pedido_Id     = " & objPedido.Codigo & vbCrLf &
                            "  and Produto_Id    = '" & ItemPedido.CodigoProduto & "'"

        Sqls.Add(Sql)

        Dim Lancamento As New [Lib].Negocio.LancamentoItemPedido(ItemPedido)

        Lancamento.IUD = "I"
        Lancamento.CodigoPedidoItem = objPedido.Itens(index).Lancamentos.Max(Function(s) s.CodigoPedidoItem) + 1
        Lancamento.TipoLancamento = eTiposLancamentosPedidos.Complemento

        Lancamento.Movimento = DateTime.Now
        Lancamento.DataEntrega = objPedido.DataEntregaFinal
        Lancamento.UsuarioLiberacao = Session("ssNomeUsuario")

        Lancamento.UnitarioOficialCompra = 0
        Lancamento.UnitarioMoedaCompra = 0

        Lancamento.QuantidadeFaturamento = 0
        Lancamento.QuantidadeComercializacao = 0

        Lancamento.UnitarioOficial = 0
        Lancamento.TotalOficial = CDec(CType(grdItens.Rows(index).FindControl("txtValorComplemento"), TextBox).Text)

        Lancamento.UnitarioMoeda = 0
        Lancamento.TotalMoeda = 0

        objPedido.Itens(index).Lancamentos.AdicionarLancamento(Lancamento)

        ItemPedido.Encargos = Nothing
        ItemPedido.Encargos.CriaListar()

        'Os valores em dolar não dever ser alterados, o sistema tem que deixar como estava no banco de dados
        Dim itemPedidoOld As PedidoXItem = New [Lib].Negocio.ListPedidoXItem(objPedido).ElementAtOrDefault(index)

        For Each enc In ItemPedido.Encargos

            Dim encOld = itemPedidoOld.Encargos.FirstOrDefault(Function(x) x.CodigoEncargo = enc.CodigoEncargo)

            If encOld IsNot Nothing Then
                enc.BaseMoeda = encOld.BaseMoeda
                enc.ValorMoeda = encOld.ValorMoeda
            Else
                enc.BaseMoeda = 0
                enc.ValorMoeda = 0
            End If

        Next

        ItemPedido.Encargos.AtualizaLiquido()

        Lancamento.SalvarSql(Sqls)
        ItemPedido.Encargos.SalvarSql(Sqls)

        If objPedido.SubOperacao.Financeiro AndAlso objPedido.Vencimentos.Count > 0 Then

            Dim objVencimentos As New Hashtable
            Dim i As Integer
            Dim temProvisao As Boolean = False

            For i = 0 To objPedido.Vencimentos.OrderBy(Function(s) s.Provisao).Count - 1

                If objPedido.Vencimentos(i).Provisao = [Lib].Negocio.eProvisao.Provisao Then
                    objVencimentos.Add(i, objPedido.Vencimentos(i).Codigo)
                    temProvisao = True
                End If
            Next

            Session("objPedVencimentos" & HID.Value) = objVencimentos

            objPedido.Vencimentos.CriaParcelaDeComplementoEmDolar(CDec(CType(grdItens.Rows(index).FindControl("txtValorComplemento"), TextBox).Text))

            objPedido.Vencimentos.ModificarHistorico(eTabelas.Pedido, New String() {objPedido.Codigo.ToString()})

            Sqls.AddRange(objPedido.Vencimentos.GetSQL(eModoAlteracao.Inclusao, UsuarioServidor.NomeServidor.ToUpper(), objPedido.MomentoFinanceiro))
        End If

        If Sqls.Count > 0 Then
            If Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, "Parcela do Complemento criada com Sucesso.", eTitulo.Sucess)
                atualizarGrid()
            Else
                MsgBox(Me.Page, "Houve um erro ao gravar as alterações, contate o suporte.", eTitulo.Erro)
            End If
        End If

    End Sub

    Private Sub confirmarEstorno(index As Integer)
        recuperaPedido()

        Dim Sqls As New ArrayList

        Dim ItemPedido As [Lib].Negocio.PedidoXItem
        ItemPedido = objPedido.Itens(index)
        ItemPedido.IUD = "I"

        Dim Sql As String = "DELETE FROM PedidosXEncargos" & vbCrLf &
                            "Where Empresa_Id    = '" & objPedido.CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf &
                            "  and Pedido_Id     = " & objPedido.Codigo & vbCrLf &
                            "  and Produto_Id    = '" & ItemPedido.CodigoProduto & "'"

        Sqls.Add(Sql)

        Dim Lancamento As New [Lib].Negocio.LancamentoItemPedido(ItemPedido)

        Lancamento.IUD = "I"
        Lancamento.CodigoPedidoItem = objPedido.Itens(index).Lancamentos.Max(Function(s) s.CodigoPedidoItem) + 1
        Lancamento.TipoLancamento = eTiposLancamentosPedidos.Estorno

        Lancamento.Movimento = DateTime.Now
        Lancamento.DataEntrega = objPedido.DataEntregaFinal
        Lancamento.UsuarioLiberacao = Session("ssNomeUsuario")


        Lancamento.UnitarioOficialCompra = 0
        Lancamento.UnitarioMoedaCompra = 0

        Lancamento.QuantidadeFaturamento = 0
        Lancamento.QuantidadeComercializacao = 0

        Lancamento.UnitarioOficial = 0
        Lancamento.TotalOficial = CDec(CType(grdItens.Rows(index).FindControl("txtValorEstorno"), TextBox).Text)

        Lancamento.UnitarioMoeda = 0
        Lancamento.TotalMoeda = 0

        objPedido.Itens(index).Lancamentos.AdicionarLancamentoEstorno(Lancamento)

        ItemPedido.Encargos = Nothing
        ItemPedido.Encargos.CriaListar()

        'Os valores em dolar não dever ser alterados, o sistema tem que deixar como estava no banco de dados
        Dim itemPedidoOld As PedidoXItem = New [Lib].Negocio.ListPedidoXItem(objPedido).ElementAtOrDefault(index)

        For Each enc In ItemPedido.Encargos

            Dim encOld = itemPedidoOld.Encargos.FirstOrDefault(Function(x) x.CodigoEncargo = enc.CodigoEncargo)

            If encOld IsNot Nothing Then
                enc.BaseMoeda = encOld.BaseMoeda
                enc.ValorMoeda = encOld.ValorMoeda
            Else
                enc.BaseMoeda = 0
                enc.ValorMoeda = 0
            End If

        Next

        ItemPedido.Encargos.AtualizaLiquido()

        Lancamento.SalvarSql(Sqls)
        ItemPedido.Encargos.SalvarSql(Sqls)

        If Sqls.Count > 0 Then
            If Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, "Estorno criada com Sucesso.", eTitulo.Sucess)
                atualizarGrid()
            Else
                MsgBox(Me.Page, "Houve um erro ao gravar as alterações, contate o suporte.", eTitulo.Erro)
            End If
        End If

    End Sub

    Private Function totalPedido()
        Dim total As Decimal

        For i = 0 To grdItens.Rows.Count - 1
            total += CDec(grdItens.Rows(i).Cells(4).Text)
        Next

        Return total
    End Function

    Protected Sub btnAlterarItem_Click(sender As Object, e As EventArgs)
        Try
            Dim txt As TextBox
            Dim btn As Button = CType(sender, Button)

            btn.Visible = False
            Dim index As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex

            btn = CType(grdItens.Rows(index).FindControl("btnSalvarItem"), Button)
            btn.Visible = True

            btn = CType(grdItens.Rows(index).FindControl("btnCancelarItem"), Button)
            btn.Visible = True

            txt = CType(grdItens.Rows(index).FindControl("txtUnitario"), TextBox)
            txt.Enabled = True

            txt = CType(grdItens.Rows(index).FindControl("txtQtd"), TextBox)
            txt.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnSalvarItem_Click(sender As Object, e As EventArgs)
        Try
            Dim txt As TextBox
            Dim btn As Button = CType(sender, Button)
            Dim index As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex

            atualizarItem(index)

            txtAtual.Text = totalPedido()

            btn.Visible = False

            btn = CType(grdItens.Rows(index).FindControl("btnCancelarItem"), Button)
            btn.Visible = False

            btn = CType(grdItens.Rows(index).FindControl("btnAlterarItem"), Button)
            btn.Visible = True

            txt = CType(grdItens.Rows(index).FindControl("txtUnitario"), TextBox)
            txt.Enabled = False

            txt = CType(grdItens.Rows(index).FindControl("txtQtd"), TextBox)
            txt.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCancelarItem_Click(sender As Object, e As EventArgs)
        Try
            Dim txt As TextBox
            Dim btn As Button = CType(sender, Button)
            Dim index As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex
            btn.Visible = False

            recuperaPedido()

            Dim quantidade As Decimal
            Dim unitario As Decimal

            For Each item In objPedido.Itens.Where(Function(s) s.CodigoProduto = grdItens.Rows(index).Cells(0).Text)
                quantidade = item.Lancamentos(0).QuantidadeFaturamento
                unitario = item.Lancamentos(0).UnitarioOficial
            Next

            CType(grdItens.Rows(index).FindControl("txtQtd"), TextBox).Text = quantidade.ToString("N4")
            CType(grdItens.Rows(index).FindControl("txtUnitario"), TextBox).Text = unitario.ToString("N10")

            btn = CType(grdItens.Rows(index).FindControl("btnSalvarItem"), Button)
            btn.Visible = False

            btn = CType(grdItens.Rows(index).FindControl("btnAlterarItem"), Button)
            btn.Visible = True

            txt = CType(grdItens.Rows(index).FindControl("txtUnitario"), TextBox)
            txt.Enabled = False

            txt = CType(grdItens.Rows(index).FindControl("txtQtd"), TextBox)
            txt.Enabled = False

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        If Not Funcoes.VerificaPermissao("AjustarPedido", "ALTERAR") Then
            MsgBox(Me.Page, "Usuário sem permissão para ajustar pedido!", eTitulo.Info)
        End If

        recuperaPedido()

        Dim Sqls As New ArrayList

        objPedido.Itens(row.RowIndex).IUD = "D"

        objPedido.Itens(row.RowIndex).SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then

            If objPedido.Vencimentos Is Nothing OrElse objPedido.Vencimentos.Count = 0 Then
                atualizarGrid()
            Else

                Sqls.Clear()

                objPedido = New Pedido(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtPedido.Text)

                If objPedido.Vencimentos IsNot Nothing AndAlso objPedido.Vencimentos.Count > 0 Then
                    Dim objVencimentos As New Hashtable
                    Dim i As Integer
                    For i = 0 To objPedido.Vencimentos.OrderBy(Function(s) s.Provisao).Count - 1
                        objVencimentos.Add(i, objPedido.Vencimentos(i).Codigo)
                    Next
                    Session("objPedVencimentos" & HID.Value) = objVencimentos
                End If

                objPedido.IUD = "U"
                objPedido.Vencimentos.CriarParcelamento(False, CType(Session("objPedVencimentos" & HID.Value), Hashtable))

                Sqls.AddRange(objPedido.Vencimentos.GetSQL(eModoAlteracao.Inclusao, UsuarioServidor.NomeServidor.ToUpper(), objPedido.MomentoFinanceiro))

                If Banco.GravaBanco(Sqls) Then
                    atualizarGrid()
                Else
                    MsgBox(Me.Page, "Houve um erro ao gravar as alterações, contate o suporte.", eTitulo.Erro)
                End If

            End If
        Else
            MsgBox(Me.Page, "Houve um erro ao gravar as alterações, contate o suporte.", eTitulo.Erro)
        End If

    End Sub

    Protected Sub btnAlterarComplemento_Click(sender As Object, e As EventArgs)
        Try
            If Funcoes.VerificaPermissao("AjustarPedido", "LIBERAR") Then

                Dim btn As Button = CType(sender, Button)
                Dim index As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex

                If CType(grdItens.Rows(index).FindControl("txtValorComplemento"), TextBox).Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor do Complemento não foi informado!", eTitulo.Info)
                ElseIf CDec(CType(grdItens.Rows(index).FindControl("txtValorComplemento"), TextBox).Text) = 0 Then
                    MsgBox(Me.Page, "Valor do Complemento não pode ser ZERO(0)!", eTitulo.Info)
                Else
                    confirmarComplemento(index)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para complementar pedido!", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAlterarEstorno_Click(sender As Object, e As EventArgs)
        Try
            If Funcoes.VerificaPermissao("AjustarPedido", "LIBERAR") Then

                Dim btn As Button = CType(sender, Button)
                Dim index As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex

                If CType(grdItens.Rows(index).FindControl("txtValorEstorno"), TextBox).Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor do Estorno não foi informado!", eTitulo.Info)
                ElseIf CDec(CType(grdItens.Rows(index).FindControl("txtValorEstorno"), TextBox).Text) = 0 Then
                    MsgBox(Me.Page, "Valor do Estorno não pode ser ZERO(0)!", eTitulo.Info)
                Else
                    confirmarEstorno(index)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para complementar pedido!", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEncargos_Click(sender As Object, e As EventArgs)
        Try
            Dim index As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex

            ucPedidoxEncargo.Limpar()
            ucPedidoxEncargo.SetarParametros(index)
            Popup.ConsultaEncargosPedido(Me.Page, "objEncargosPedido" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If txtPedido.Text.Length = 0 Then
                MsgBox(Me.Page, "Número do Pedido não foi informado!", eTitulo.Info)
            ElseIf ddlEmpresa.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada!", eTitulo.Info)
            Else
                atualizarGrid()
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            If Funcoes.VerificaPermissao("AjustarPedido", "ALTERAR") Then
                confirmarAtualizacoes()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para ajustar pedido!", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AjustarPedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub cmdConsultarCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultarCliente.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4,5")
            Popup.ConsultaDeClientes(Me, "objClientePXI" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class