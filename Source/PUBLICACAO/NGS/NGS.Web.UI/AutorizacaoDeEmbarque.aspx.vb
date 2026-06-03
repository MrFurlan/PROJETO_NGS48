Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AutorizacaoDeEmbarque
    Inherits BasePage

#Region "Variáveis Locais"
    Private ListEmbarquePedido As [Lib].Negocio.ListEmbarquePedido
    Private objEmbarquePedido As [Lib].Negocio.EmbarquePedido
#End Region

#Region "Sessão"
    Private Sub SessaoSalvarListEmbarquePedido()
        Session("ssListEmbarque" & HID.Value) = ListEmbarquePedido
    End Sub

    Private Sub SessaoRecuperaListEmbarquePedido()
        ListEmbarquePedido = CType(Session("ssListEmbarque" & HID.Value), [Lib].Negocio.ListEmbarquePedido)
    End Sub

    Private Sub SessaoSalvarObjEmbarque()
        Session("ssEmbarquePedido" & HID.Value) = objEmbarquePedido
    End Sub

    Private Sub SessaoRecuperaObjEmbarque()
        objEmbarquePedido = CType(Session("ssEmbarquePedido" & HID.Value), [Lib].Negocio.EmbarquePedido)
    End Sub
#End Region

#Region "Métodos"
    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        Try
            If Session("objClienteAutEmb" & HID.Value) IsNot Nothing Then
                Dim objCliente As Cliente = Session("objClienteAutEmb" & HID.Value)
                txtNomeCliente.Text = Funcoes.FormatarListItemCliente(objCliente).Text
                txtCodigoCliente.Value = Funcoes.FormatarListItemCliente(objCliente).Value
                Session.Remove("objClienteAutEmb" & HID.Value)

            ElseIf Session("objClienteEntregaAutEmb" & HID.Value) IsNot Nothing Then
                SessaoRecuperaObjEmbarque()
                Dim objCliente As Cliente = Session("objClienteEntregaAutEmb" & HID.Value)
                Dim pos As Integer = objEmbarquePedido.LocaisDeEntrega.FindIndex(Function(s) s.CodigoClienteEntrega = objCliente.Codigo And s.EndClienteEntrega = objCliente.CodigoEndereco)

                Funcoes.FormatarClienteTXT(txtNomeLocaldeEntrega, objCliente)
                txtCodigoLocalEntrega.Value = objCliente.Codigo & "-" & objCliente.CodigoEndereco

                If pos > -1 Then
                    MsgBox(Me.Page, "Local de Entrega Já Adicionado!")
                    chkContaEOrdem.Checked = objEmbarquePedido.LocaisDeEntrega(pos).EmitirNota
                    Exit Sub
                End If

                chkContaEOrdem.Checked = False
                Session.Remove("objClienteEntregaAutEmb" & HID.Value)

            ElseIf Session("objProdutoAutEmb" & HID.Value) IsNot Nothing Then
                Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoAutEmb" & HID.Value)
                txtNomeProduto.Text = objProduto.Descricao
                CodigoProduto.Value = objProduto.Codigo
                Session.Remove("objProdutoAutEmb" & HID.Value)

            ElseIf Session("TransportadorRoteiro" & HID.Value) IsNot Nothing Then
                SessaoRecuperaObjEmbarque()
                Dim posEntrega As Integer = gridLocalEntrega.SelectedIndex
                Dim posRoteiro As Integer = grdRoteiro.SelectedIndex

                Dim objCliente As Cliente = Session("TransportadorRoteiro" & HID.Value)
                Dim pos As Integer = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores.FindIndex(Function(s) s.CodigoTransportador = objCliente.Codigo And s.EndTransportador = objCliente.CodigoEndereco)

                Funcoes.FormatarClienteTXT(txtNomeTransportador, objCliente)
                txtCodigoTransportador.Value = objCliente.Codigo & "-" & objCliente.CodigoEndereco

                If pos > -1 Then
                    MsgBox(Me.Page, "Transportador já Incluido!")
                    txtQuota.Text = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(pos).Quota.ToString("N2")
                    Exit Sub
                End If

                Session.Remove("objClienteAutEmb" & HID.Value)
                SessaoSalvarObjEmbarque()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal str As String)
        Try
            If grdRoteiro IsNot Nothing OrElse grdRoteiro.Rows.Count = 0 Then
                SessaoRecuperaObjEmbarque()
                Dim posEntrega As Integer = gridLocalEntrega.SelectedIndex
                Dim cli As String() = str.Split(";")
                Dim origem As String() = cli(0).Split("-")
                Dim destino As String() = cli(1).Split("-")
                Dim roteiro As New EmbarqueRoteiro(objEmbarquePedido.LocaisDeEntrega(posEntrega))

                If Not ExisteRoteiro(cli, posEntrega) Then
                    roteiro.IUD = "I"
                    roteiro.CodigoOrigem = origem(0)
                    roteiro.EndOrigem = origem(1)
                    roteiro.CodigoDestino = destino(0)
                    roteiro.EndDestino = destino(1)
                    roteiro.CodigoViaDeTransporte = cli(2)

                    If roteiro.Salvar() Then
                        roteiro.IUD = ""
                        objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros.Add(roteiro)
                    End If

                    grdRoteiro.DataSource = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros
                    grdRoteiro.DataBind()
                    SessaoSalvarObjEmbarque()
                Else
                    MsgBox(Me.Page, "Roteiro já existe no pedido - " & objEmbarquePedido.CodigoPedido & ".")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ExisteRoteiro(ByVal Clientes() As String, ByVal posEntrega As Integer) As Boolean
        Dim origem As String = Clientes(0)
        Dim destino As String = Clientes(1)
        For Each obj As EmbarqueRoteiro In objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros

            If obj.ParentEntrega.ParentEmbPedido.CodigoEmpresa = objEmbarquePedido.CodigoEmpresa _
                AndAlso obj.ParentEntrega.ParentEmbPedido.CodigoPedido = objEmbarquePedido.CodigoPedido _
                AndAlso obj.CodigoOrigem & "-" & obj.EndOrigem = origem _
                AndAlso obj.CodigoDestino & "-" & obj.EndDestino = destino Then Return True
        Next
        Return False
    End Function

    Public Sub CarregarResumoEmbarque()
        SessaoRecuperaObjEmbarque()
        gridLocalEntrega.DataSource = objEmbarquePedido.LocaisDeEntrega.ToArray
        gridLocalEntrega.DataBind()
    End Sub

    Private Sub Limpar()
        gridLocalEntrega.DataSource = New List(Of Object)
        gridLocalEntrega.DataBind()
        grdRoteiro.DataSource = New List(Of Object)
        grdRoteiro.DataBind()
        grdTransportador.DataSource = New List(Of Object)
        grdTransportador.DataBind()
        grdPrecos.DataSource = New List(Of Object)
        grdPrecos.DataBind()
        'pnlRoteiro.Visible = False
        pnlTransportador.Visible = False
        pnlPreco.Visible = False
        txtDataEntregaAut.Text = String.Empty
        txtPedidoAut.Text = String.Empty
        txtPedido.Text = String.Empty
        txtSafraAut.Text = String.Empty
        txtCodigoClienteAut.Value = String.Empty
        txtClienteAut.Text = String.Empty
        txtEmpresaAut.Text = String.Empty
        txtCodigoEmpresaAut.Value = String.Empty
        ddlEmpresa.SelectedIndex = 0
        txtNomeCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtNomeProduto.Text = String.Empty
        CodigoProduto.Value = String.Empty
        ddlSafra.SelectedIndex = 0
        rdTodos.Checked = True

        txtCfop.Text = String.Empty
        txtOperacao.Text = String.Empty
        txtProduto.Text = String.Empty

        txtNomeLocaldeEntrega.Text = String.Empty
        txtCodigoLocalEntrega.Value = String.Empty
        lblRecebeLocalDeEntrega.Text = String.Empty

        txtNomeTransportador.Text = String.Empty
        txtQuota.Text = String.Empty
        txtCodigoTransportador.Value = String.Empty

        grdPedidos.DataSource = New List(Of Object)
        grdPedidos.DataBind()

        Session.Remove("ssEmbarque" & HID.Value)
        TabContainer1.ActiveTabIndex = 0
    End Sub

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AUTORIZACAODERETIRADA", "ACESSAR") Then
                HID.Value = Guid.NewGuid.ToString()
                ucAutorizacaoDeEmbarque.SetarHID(HID.Value)
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaProduto.SetarHID(HID.Value)
                ucOrigemDestino.SetarHID(HID.Value)
                ucConsultaObservacoes.SetarHID(HID.Value)
                ucConsultaObservacoesEmbarque.SetarHID(HID.Value)

                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas)
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteAutEmb" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProduto_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnProduto.Click
        Try
            Session("Where" & HID.Value) = " Situacao = 1 "
            ucConsultaProduto.Limpar()
            ucConsultaProduto.BuscarProduto()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoAutEmb" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultarNovo.Click
        Try
            If String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) AndAlso String.IsNullOrWhiteSpace(txtPedido.Text) Then
                MsgBox(Me.Page, "Informe a safra ou o pedido.")
            Else
                Dim EntradaSaida As String = ""
                Dim empresa As String = ""
                Dim endEmpresa As String = ""
                Dim cliente As String = ""
                Dim endCliente As String = ""

                If ddlEmpresa.SelectedIndex > 0 Then
                    empresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    endEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                End If

                If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                    cliente = txtCodigoCliente.Value.Split("-")(0)
                    endCliente = txtCodigoCliente.Value.Split("-")(1)
                End If

                If rdEntrada.Checked Then
                    EntradaSaida = "E"
                ElseIf rdSaida.Checked Then
                    EntradaSaida = "S"
                End If

                ListEmbarquePedido = New ListEmbarquePedido(txtPedido.Text.Trim(), ddlSafra.SelectedValue, empresa, endEmpresa, cliente, endCliente, EntradaSaida)
                grdPedidos.DataSource = ListEmbarquePedido.ToArray()
                grdPedidos.DataBind()
                SessaoSalvarListEmbarquePedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdPedidos_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdPedidos.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.DataItem IsNot Nothing Then
                Try
                    Dim btn As Button = CType(e.Row.FindControl("BtnAlteraSituacaoPedido"), Button)
                    If CBool(btn.Text) Then
                        btn.Text = "Ativo"
                        'btn.BackColor = Drawing.Color.Green
                        btn.BorderColor = Drawing.Color.Green
                        btn.ForeColor = Drawing.Color.Green
                    Else
                        btn.Text = "Bloqueado"
                        'btn.BackColor = Drawing.Color.Red
                        btn.BorderColor = Drawing.Color.Red
                        btn.ForeColor = Drawing.Color.Red
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            End If
        End If
    End Sub

    Protected Sub BtnAlteraSituacaoPedido_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SessaoRecuperaListEmbarquePedido()
            Dim btn As Button = CType(sender, Button)
            Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)
            Dim situacao As Integer = 0

            If btn.Text = "Ativo" Then
                btn.Text = "Bloqueado"
                'btn.BackColor = Drawing.Color.Red
                btn.BorderColor = Drawing.Color.Red
                btn.ForeColor = Drawing.Color.Red
            Else
                btn.Text = "Ativo"
                'btn.BackColor = Drawing.Color.Green
                btn.BorderColor = Drawing.Color.Green
                btn.ForeColor = Drawing.Color.Green
                situacao = 1
            End If

            Dim objPedido As Pedido = ListEmbarquePedido(row.RowIndex).Pedido
            If objPedido IsNot Nothing Then
                objPedido.IUD = "U"
                objPedido.CodigoSituacao = situacao
                If objPedido.AtualizaEmbarqueAtivo(situacao) Then objPedido.IUD = ""
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdPedidos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPedidos.SelectedIndexChanged
        Try
            SessaoRecuperaListEmbarquePedido()
            objEmbarquePedido = ListEmbarquePedido(grdPedidos.SelectedIndex)
            SessaoSalvarObjEmbarque()

            txtPedidoAut.Text = objEmbarquePedido.CodigoPedido
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objEmbarquePedido.Cliente)
            txtClienteAut.Text = itemCliente.Text
            txtCodigoClienteAut.Value = itemCliente.Value
            txtDataEntregaAut.Text = objEmbarquePedido.DataEntrega.ToString("dd/MM/yyyy")
            txtProduto.Text = objEmbarquePedido.CodigoProduto & " - " & objEmbarquePedido.NomeProduto
            txtOperacao.Text = objEmbarquePedido.DescOperacao
            lblUnidade.Text = objEmbarquePedido.Produto.Unidade
            ddlPesoQuantidade.SelectedValue = "P"

            Dim objCfop As [Lib].Negocio.CFOP = New [Lib].Negocio.CFOP(objEmbarquePedido.Pedido.Itens(0).OperacaoxEstado.CodigoFiscal)
            txtCfop.Text = objCfop.Codigo.ToString + " - " + objCfop.Descricao

            chkContaEOrdem.Checked = False
            chkContaEOrdem.Visible = objEmbarquePedido.Pedido.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString

            gridLocalEntrega.DataSource = objEmbarquePedido.LocaisDeEntrega.ToArray
            gridLocalEntrega.DataBind()

            txtSafraAut.Text = objEmbarquePedido.CodigoSafra
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmbarquePedido.Empresa)
            txtCodigoEmpresaAut.Value = itemEmpresa.Value
            txtEmpresaAut.Text = itemEmpresa.Text

            If objEmbarquePedido.LocaisDeEntrega.Count = 0 Then
                txtNomeLocaldeEntrega.Text = objEmbarquePedido.ClienteFormatado
                chkContaEOrdem.Checked = False
                txtCodigoLocalEntrega.Value = objEmbarquePedido.CodigoCliente & "-" & objEmbarquePedido.EndCliente
            Else
                txtNomeLocaldeEntrega.Text = ""
                chkContaEOrdem.Checked = False
                txtCodigoLocalEntrega.Value = ""
            End If

            TabContainer1.ActiveTabIndex = 1
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovoLocalEntregaClientePedido_Click(sender As Object, e As EventArgs) Handles lnkNovoLocalEntregaClientePedido.Click
        Try
            If String.IsNullOrWhiteSpace(txtCodigoLocalEntrega.Value) Then
                MsgBox(Me.Page, "Informe o Local de entrega.")
            Else
                SessaoRecuperaObjEmbarque()
                Dim cli As String() = txtCodigoLocalEntrega.Value.Split("-")
                Dim pos As Integer = objEmbarquePedido.LocaisDeEntrega.FindIndex(Function(s) s.CodigoClienteEntrega = cli(0) AndAlso s.EndClienteEntrega = cli(1))

                If pos > -1 Then
                    If objEmbarquePedido.LocaisDeEntrega(pos).EmitirNota = chkContaEOrdem.Checked OrElse (objEmbarquePedido.CodigoCliente = cli(0) AndAlso objEmbarquePedido.EndCliente = cli(1)) Then
                        MsgBox(Me.Page, "Cliente do Pedido já Adicionado.")
                        Exit Sub
                    End If

                    'objEmbarquePedido.LocaisDeEntrega(pos).EmitirNota = chkContaEOrdem.Checked
                    'objEmbarquePedido.LocaisDeEntrega(pos).IUD = "U"
                    'objEmbarquePedido.LocaisDeEntrega(pos).Salvar()
                    'gridLocalEntrega.DataSource = objEmbarquePedido.LocaisDeEntrega.ToArray
                    'gridLocalEntrega.DataBind()
                    'SessaoSalvarObjEmbarque()
                    'Exit Sub
                ElseIf objEmbarquePedido.CodigoCliente = cli(0) AndAlso objEmbarquePedido.EndCliente = cli(1) Then
                    chkContaEOrdem.Checked = False
                End If

                Dim objLocalEntrega As New EmbarqueXEntrega(objEmbarquePedido)
                objLocalEntrega.IUD = "I"
                objLocalEntrega.CodigoClienteEntrega = cli(0)
                objLocalEntrega.EndClienteEntrega = cli(1)
                objLocalEntrega.EmitirNota = chkContaEOrdem.Checked

                If objLocalEntrega.Salvar() Then
                    objLocalEntrega.IUD = ""
                    objEmbarquePedido.LocaisDeEntrega.Add(objLocalEntrega)
                    SessaoSalvarObjEmbarque()

                    gridLocalEntrega.DataSource = objEmbarquePedido.LocaisDeEntrega.ToArray
                    gridLocalEntrega.DataBind()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaLocalEntrega_Click(sender As Object, e As EventArgs) Handles btnConsultaLocalEntrega.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteEntregaAutEmb" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridLocalEntrega_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridLocalEntrega.SelectedIndexChanged
        Try
            SessaoRecuperaObjEmbarque()
            Dim Linha As Integer = gridLocalEntrega.SelectedIndex

            Funcoes.FormatarClienteTXT(txtNomeLocaldeEntrega, objEmbarquePedido.LocaisDeEntrega(Linha).ClienteEntrega)
            txtCodigoLocalEntrega.Value = objEmbarquePedido.LocaisDeEntrega(Linha).CodigoClienteEntrega + "-" + objEmbarquePedido.LocaisDeEntrega(Linha).EndClienteEntrega
            chkContaEOrdem.Checked = objEmbarquePedido.LocaisDeEntrega(Linha).EmitirNota

            lblRecebeLocalDeEntrega.Text = "LOCAL ENTREGA: " + txtNomeLocaldeEntrega.Text

            grdRoteiro.DataSource = objEmbarquePedido.LocaisDeEntrega(Linha).Roteiros
            grdRoteiro.DataBind()

            'pnlRoteiro.Visible = True
            pnlTransportador.Visible = False
            pnlPreco.Visible = False
            SessaoSalvarObjEmbarque()

            TabContConfigPedido.ActiveTabIndex = 1
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAutEmbarque_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim row As GridViewRow = CType(CType(sender, LinkButton).NamingContainer, GridViewRow)

            SessaoRecuperaObjEmbarque()
            ucAutorizacaoDeEmbarque.SetarLocalEntregaProduto(row.RowIndex, 0)
            Popup.ConsultaDeAutorizacaoDeEmbarque(Me.Page, "objAutEmbarque" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirAutorizacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaObjEmbarque()
            Dim row As GridViewRow = CType(CType(sender, ImageButton).NamingContainer, GridViewRow)

            objEmbarquePedido.LocaisDeEntrega(row.RowIndex).IUD = "D"
            If objEmbarquePedido.LocaisDeEntrega(row.RowIndex).Salvar Then
                objEmbarquePedido.LocaisDeEntrega.RemoveAt(row.RowIndex)
                MsgBox(Me.Page, "Registro(s) excluído com Sucesso.", eTitulo.Sucess)
            Else
                objEmbarquePedido.LocaisDeEntrega(row.RowIndex).IUD = ""
            End If

            gridLocalEntrega.DataSource = objEmbarquePedido.LocaisDeEntrega
            grdPrecos.DataBind()

            SessaoSalvarObjEmbarque()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdicionarObservacao_Click(sender As Object, e As EventArgs)
        Try
            Dim row As GridViewRow = CType(CType(sender, LinkButton).NamingContainer, GridViewRow)

            SessaoRecuperaObjEmbarque()
            ucConsultaObservacoesEmbarque.SetarLocalEntregaProduto(row.RowIndex)
            Popup.ConsultaDeObservacoesEmbarque(Me.Page, "objObservacaoEmbarque" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdicionarRoteiro_Click(sender As Object, e As EventArgs) Handles lnkAdicionarRoteiro.Click
        Try
            ucOrigemDestino.Limpar()
            Popup.CadastrarOrigemDestinoRoteiro(Me.Page, "OrigemDestinoRoteiro" & HID.Value, "")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirRoteiro_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaObjEmbarque()
            Dim row As GridViewRow = CType(CType(sender, ImageButton).NamingContainer, GridViewRow)

            objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(row.RowIndex).IUD = "D"
            If objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(row.RowIndex).Salvar() Then
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros.RemoveAt(row.RowIndex)
                MsgBox(Me.Page, "Registro(s) excluído com Sucesso.", eTitulo.Sucess)
            Else
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(row.RowIndex).IUD = ""
            End If

            grdRoteiro.DataSource = objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros
            grdRoteiro.DataBind()

            pnlTransportador.Visible = False
            pnlPreco.Visible = False

            SessaoSalvarObjEmbarque()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridRoteiro_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdRoteiro.SelectedIndexChanged
        Try
            SessaoRecuperaObjEmbarque()

            Dim posEntrega As Integer = gridLocalEntrega.SelectedIndex
            Dim posRoteiro As Integer = grdRoteiro.SelectedIndex

            lblRoteiroDe.Text = String.Format("De:   {0} ({1}) / {2}-{3}", objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Origem.Nome, Funcoes.FormatarCpfCnpj(objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).CodigoOrigem), objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Origem.Cidade, objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Origem.CodigoEstado)
            lblRoteiroPara.Text = String.Format("Para: {0} ({1}) / {2}-{3}", objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Destino.Nome, Funcoes.FormatarCpfCnpj(objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).CodigoDestino), objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Destino.Cidade, objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Destino.CodigoEstado)

            pnlTransportador.Visible = True
            pnlPreco.Visible = False

            grdTransportador.DataSource = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores
            grdTransportador.DataBind()
            SessaoSalvarObjEmbarque()

            TabContConfigPedido.ActiveTabIndex = 2
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovoTransportador_Click(sender As Object, e As EventArgs) Handles lnkNovoTransportador.Click
        Try
            If String.IsNullOrWhiteSpace(txtCodigoTransportador.Value) Then
                MsgBox(Me.Page, "Informe o transportador.")
            ElseIf Not IsNumeric(txtQuota.Text) OrElse CDec(txtQuota.Text) = 0 Then
                MsgBox(Me.Page, "Informe o valor para quota a ser negociada.")
            Else
                SessaoRecuperaObjEmbarque()
                If Not IsNumeric(txtQuota.Text) Then txtQuota.Text = "0,0000"
                Dim posEntrega As Integer = gridLocalEntrega.SelectedIndex
                Dim posRoteiro As Integer = grdRoteiro.SelectedIndex
                Dim Cli As String() = txtCodigoTransportador.Value.Split("-")

                Dim pos As Integer = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores.FindIndex(Function(s) s.CodigoTransportador = Cli(0) And s.EndTransportador = Cli(1))

                If pos > -1 Then
                    If CDec(objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(pos).Quota) = CDec(txtQuota.Text) Then
                        MsgBox(Me.Page, "Transportador ja Adicionado ao Roteiro.")
                        Exit Sub
                    End If
                End If

                Dim transportador As New EmbarqueTransportador(objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro))

                transportador.IUD = "I"
                transportador.CodigoTransportador = Cli(0)
                transportador.EndTransportador = Cli(1)
                transportador.Quota = CDec(txtQuota.Text)
                transportador.PesoQuantidade = ddlPesoQuantidade.SelectedValue
                If transportador.Salvar() Then
                    transportador.IUD = ""
                    objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores.Add(transportador)
                End If

                grdTransportador.DataSource = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores.ToArray
                grdTransportador.DataBind()
                SessaoSalvarObjEmbarque()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaTransportador_Click(sender As Object, e As EventArgs) Handles btnConsultaTransportador.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
            Popup.ConsultaDeClientes(Me.Page, "TransportadorRoteiro" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlPesoQuantidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPesoQuantidade.SelectedIndexChanged
        Try
            If ddlPesoQuantidade.SelectedValue = "Q" Then
                lblUnidade.Text = "FRETES"
            Else
                SessaoRecuperaObjEmbarque()
                lblUnidade.Text = objEmbarquePedido.Produto.Unidade
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdTransportador_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdTransportador.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.DataItem IsNot Nothing Then
                Try
                    Dim btn As Button = CType(e.Row.FindControl("btnAlteraAtivo"), Button)
                    btn.Text = IIf(CBool(btn.Text), "Ativo", "Bloqueado")
                    formataBotaoStatusTransportador(btn)
                Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
            End If
        End If
    End Sub

    Private Sub formataBotaoStatusTransportador(ByVal btn As Button)
        If btn.Text.Equals("Ativo") Then
            btn.Text = "Ativo"
            btn.ForeColor = Drawing.Color.Green
            btn.BorderColor = Drawing.Color.Green
        Else
            btn.Text = "Bloqueado"
            btn.BorderColor = Drawing.Color.Red
            btn.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Protected Sub btnAlteraAtivo_Click(sender As Object, e As EventArgs)
        Try
            SessaoRecuperaObjEmbarque()
            Dim btnAtivo As Button = CType(sender, Button)
            If btnAtivo IsNot Nothing Then
                Dim row As GridViewRow = CType(btnAtivo.NamingContainer, GridViewRow)
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).IUD = "U"
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).Ativo = IIf(Not btnAtivo.Text.Equals("Ativo"), True, False)

                If objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).Salvar() Then
                    objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).IUD = ""
                    SessaoSalvarObjEmbarque()
                    btnAtivo.Text = IIf(Not (btnAtivo.Text.Equals("Ativo")), "Ativo", "Bloqueado")
                    formataBotaoStatusTransportador(btnAtivo)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirTransportador_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SessaoRecuperaObjEmbarque()
            Dim row As GridViewRow = CType(CType(sender, ImageButton).NamingContainer, GridViewRow)

            objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).IUD = "D"
            If objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).Salvar() Then
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores.RemoveAt(row.RowIndex)
                MsgBox(Me.Page, "Registro(s) excluído com Sucesso.", eTitulo.Sucess)
            Else
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(row.RowIndex).IUD = ""
            End If

            grdTransportador.DataSource = objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores
            grdTransportador.DataBind()

            pnlPreco.Visible = False
            SessaoSalvarObjEmbarque()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridTransportador_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdTransportador.SelectedIndexChanged
        Try
            SessaoRecuperaObjEmbarque()
            pnlPreco.Visible = True
            Dim posEntrega As Integer = gridLocalEntrega.SelectedIndex
            Dim posRoteiro As Integer = grdRoteiro.SelectedIndex
            Dim posTransportador As Integer = grdTransportador.SelectedIndex

            lblRecebeTransp.Text = String.Format("{0} ({1}))", objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Transportador.Nome, _
                                                 Funcoes.FormatarCpfCnpj(objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).CnpjTransp))

            Funcoes.FormatarClienteTXT(txtNomeTransportador, objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Transportador)
            txtQuota.Text = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Quota.ToString("N2")
            txtCodigoTransportador.Value = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).CodigoTransportador & "-" & objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).EndTransportador
            ddlPesoQuantidade.SelectedValue = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).PesoQuantidade

            'If objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).PesoQuantidade = "P" Then
            '    lblUnidadePrecoTransportador.Text = "Por Ton."
            'Else
            '    lblUnidadePrecoTransportador.Text = "Por Frete"
            'End If

            grdPrecos.DataSource = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Precos.Where(Function(s) s.IUD <> "D")
            grdPrecos.DataBind()
            SessaoSalvarObjEmbarque()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkGravarValores_Click(sender As Object, e As EventArgs) Handles lnkGravarValores.Click
        Try
            If Not IsNumeric(txtQuotaPrecoTransportador.Text) OrElse CDec(txtQuotaPrecoTransportador.Text) = 0 _
                OrElse Not IsNumeric(txtValorPrecoTransportador.Text) OrElse CDec(txtValorPrecoTransportador.Text) = 0 Then
                MsgBox(Me.Page, "Para incluir uma cotação, informe os valores de Cota e Preço.")
                Exit Sub
            End If
            SessaoRecuperaObjEmbarque()

            Dim posEntrega As Integer = gridLocalEntrega.SelectedIndex
            Dim posRoteiro As Integer = grdRoteiro.SelectedIndex
            Dim posTransportador As Integer = grdTransportador.SelectedIndex

            Dim objPreco As New EmbarquePrecoFrete(objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador))

            objPreco.IUD = "I"
            objPreco.Movimento = DateTime.Now

            If objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Precos IsNot Nothing AndAlso objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Precos.Count > 0 Then
                objPreco.NrCotacao = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Precos.Max(Function(s) s.NrCotacao) + 1
            Else
                objPreco.NrCotacao = 1
            End If

            If Not IsNumeric(txtQuotaPrecoTransportador.Text) Then txtQuotaPrecoTransportador.Text = "0"
            objPreco.Quota = CDec(txtQuotaPrecoTransportador.Text)

            If objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).PesoQuantidade = "P" Then
                objPreco.ValorTon = CDec(txtValorPrecoTransportador.Text)
                objPreco.ValorFrete = 0
            Else
                objPreco.ValorFrete = CDec(txtValorPrecoTransportador.Text)
                objPreco.ValorTon = 0
            End If
            objPreco.UsuarioInclusao = UsuarioServidor.NomeUsuario

            If objPreco.Salvar() Then
                objPreco.IUD = ""
                objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Precos.Add(objPreco)
            End If

            grdPrecos.DataSource = objEmbarquePedido.LocaisDeEntrega(posEntrega).Roteiros(posRoteiro).Transportadores(posTransportador).Precos
            grdPrecos.DataBind()

            SessaoSalvarObjEmbarque()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirPreco_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SessaoRecuperaObjEmbarque()
            Dim row As GridViewRow = CType(CType(sender, ImageButton).NamingContainer, GridViewRow)

            objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(grdTransportador.SelectedIndex).Precos(row.RowIndex).IUD = "D"
            If objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(grdTransportador.SelectedIndex).Precos(row.RowIndex).Salvar Then
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(grdTransportador.SelectedIndex).Precos.RemoveAt(row.RowIndex)
                MsgBox(Me.Page, "Registro(s) excluído com Sucesso.", eTitulo.Sucess)
            Else
                objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(grdTransportador.SelectedIndex).Precos(row.RowIndex).IUD = ""
            End If

            grdPrecos.DataSource = objEmbarquePedido.LocaisDeEntrega(gridLocalEntrega.SelectedIndex).Roteiros(grdRoteiro.SelectedIndex).Transportadores(grdTransportador.SelectedIndex).Precos
            grdPrecos.DataBind()

            SessaoSalvarObjEmbarque()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("AutorizacaoDeEmbarque", "RELATORIO") Then
                Dim ds As DataSet = GetDataSet()

                Funcoes.BindReport(Me.Page, ds, "Cr_AutorizacaoDeEmbarque", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        'If Not String.IsNullOrWhiteSpace(DdlGrupo.sele) Then
        '    param &= "Código: " & DdlGrupo.Text
        'End If
        'If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
        '    param &= "Descrição: " & txtDescricao.Text
        'End If
        'param &= IIf(radSim.Checked, "Adiantamento: Sim", "Adiantamento: Não")

        Return param
    End Function

    Private Function GetDataSet() As DataSet
        Dim ds As DataSet
        Dim sql As String
        sql = "Select c.Cliente_Id + ' - ' + c.Nome as ClienteEntrega, ae.Produto_Id +  ' - ' +p.Nome as Produto," & vbCrLf & _
              " 	  Case" & vbCrLf & _
              " 	  	when ae.TipoDeLancamento = 'C' Then 'Complemento'" & vbCrLf & _
              " 	  	When ae.TipoDeLancamento = 'N' Then 'Normal'" & vbCrLf & _
              " 	  	Else  'Estorno'" & vbCrLf & _
              " 	  End as TipoDeLancamento," & vbCrLf & _
              "       ae.Movimento," & vbCrLf & _
              "       ae.Quantidade" & vbCrLf & _
              "  from AutEmbarque ae" & vbCrLf & _
              " Inner Join Produtos p" & vbCrLf & _
              "    On p.Produto_Id = ae.Produto_Id" & vbCrLf & _
              " Inner Join Clientes c" & vbCrLf & _
              "    On c.Cliente_Id  = ae.Entrega_Id" & vbCrLf & _
              "   And c.Endereco_Id = ae.EndEntrega_Id" & vbCrLf & _
              " WHERE ae.Empresa_Id    ='" & txtCodigoEmpresaAut.Value.Split("-")(0) & "'" & vbCrLf & _
              "   AND ae.EndEmpresa_Id = " & txtCodigoEmpresaAut.Value.Split("-")(1) & vbCrLf & _
              "   AND ae.Pedido_Id     = " & txtPedidoAut.Text & vbCrLf

        sql &= "Select aepf.Pedido_Id," & vbCrLf & _
               " 	   CASE" & vbCrLf & _
               "         WHEN aepf.ValorFrete > 0" & vbCrLf & _
               "           THEN aepf.ValorFrete" & vbCrLf & _
               "           ELSE aepf.ValorTon" & vbCrLf & _
               "       END as Preco," & vbCrLf & _
               "       CASE" & vbCrLf & _
               "         WHEN aepf.ValorFrete > 0" & vbCrLf & _
               "           THEN 'Frete'" & vbCrLf & _
               "           ELSE 'Tonelada'" & vbCrLf & _
               "       END as CobradoPor," & vbCrLf & _
               "       aepf.Quota," & vbCrLf & _
               " 	   aepf.Movimento," & vbCrLf & _
               " 	   aepf.Transportador_Id + '-' + cast(aepf.EndTransportador_Id as nvarchar) + ' - ' + cli.Nome as Transportador," & vbCrLf & _
               " 	   aer.Roteiro_Id," & vbCrLf & _
               " 	   'Origem: ' + ori.Cliente_Id + '-' + cast(ori.Endereco_Id as nvarchar) + ' - ' + ori.Nome + ' Destino: ' +  dst.Cliente_Id + ' - ' + cast(dst.Endereco_Id as nvarchar) + ' - ' + dst.Nome as Origem," & vbCrLf & _
               " 	   dst.Cliente_Id + '-' + cast(dst.Endereco_Id as nvarchar) + ' - ' + dst.Nome as Destino," & vbCrLf & _
               " 	   ent.Cliente_Id + '-' + cast(ent.Endereco_Id as nvarchar) + ' - ' + ent.Nome as Entrega" & vbCrLf & _
               "  FROM AutEmbarqueXPrecoFrete aepf" & vbCrLf & _
               " INNER JOIN Clientes cli" & vbCrLf & _
               " 	ON cli.Cliente_Id = aepf.Transportador_Id" & vbCrLf & _
               "   and cli.Endereco_Id = aepf.EndTransportador_Id" & vbCrLf & _
               " INNER JOIN AutEmbarqueXRoteiro aer" & vbCrLf & _
               " 	ON aer.Empresa_Id    = aepf.Empresa_Id" & vbCrLf & _
               "   and aer.EndEmpresa_Id = aepf.EndEmpresa_Id" & vbCrLf & _
               "   and aer.Pedido_Id  = aepf.Pedido_Id" & vbCrLf & _
               "   and aer.Roteiro_Id = aepf.Roteiro_Id" & vbCrLf & _
               "  LEFT JOIN Clientes ori" & vbCrLf & _
               "    ON ori.Cliente_Id  = aer.Origem" & vbCrLf & _
               "   and ori.Endereco_Id = aer.EndOrigem" & vbCrLf & _
               "  LEFT JOIN Clientes dst" & vbCrLf & _
               "    ON dst.Cliente_Id  = aer.Destino" & vbCrLf & _
               "   and dst.Endereco_Id = aer.EndDestino" & vbCrLf & _
               "  LEFT JOIN Clientes ent" & vbCrLf & _
               "    ON ent.Cliente_Id  = aer.Entrega" & vbCrLf & _
               "   and ent.Endereco_Id = aer.EndEntrega" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigoEmpresaAut.Value) Then
            sql &= "   AND aer.Empresa_Id    ='" & txtCodigoEmpresaAut.Value.Split("-")(0) & "'" & vbCrLf & _
                   "   AND aer.EndEmpresa_Id = " & txtCodigoEmpresaAut.Value.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedidoAut.Text) Then
            sql &= "   AND aer.Pedido_Id =  " & txtPedidoAut.Text & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "AutorizacaoDeEmbarque")
        ds.Tables(0).TableName = "Lancamentos"
        ds.Tables(1).TableName = "AutorizacaoDeEmbarque"

        Return ds
    End Function

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AUTORIZACAODERETIRADA")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class