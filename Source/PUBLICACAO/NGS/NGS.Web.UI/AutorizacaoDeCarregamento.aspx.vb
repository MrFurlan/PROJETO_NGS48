Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AutorizacaoDeCarregamento
    Inherits BasePage

#Region "Variáveis Locais"

    Private ListEmbarquePedido As [Lib].Negocio.ListEmbarquePedido
    Private objEmbarquePedido As [Lib].Negocio.EmbarquePedido
    Private objCarregamento As [Lib].Negocio.AutCarregamento
    Private listCarregamentoXNotas As [Lib].Negocio.ListAutCarregamentoXNotaFiscal

#End Region

#Region "Métodos"

    Private Sub SalvarSessaoAutCarregamento()
        Session("objCarregamento" & HID.Value) = objCarregamento
    End Sub

    Private Sub RecuperarSessaoAutCarregamento()
        objCarregamento = CType(Session("objCarregamento" & HID.Value), [Lib].Negocio.AutCarregamento)
    End Sub

    Private Sub SessaoSalvarListEmbarquePedido()
        Session("ssListEmbarque" & HID.Value) = ListEmbarquePedido
    End Sub

    Private Sub SessaoRecuperaListEmbarquePedido()
        ListEmbarquePedido = CType(Session("ssListEmbarque" & HID.Value), [Lib].Negocio.ListEmbarquePedido)
    End Sub

    Private Sub SessaoSalvarListCarregamentoXNotas()
        Session("ssListCarregamentoXNotas" & HID.Value) = listCarregamentoXNotas
    End Sub

    Private Sub SessaoRecuperaListCarregamentoXNotas()
        listCarregamentoXNotas = CType(Session("ssListCarregamentoXNotas" & HID.Value), [Lib].Negocio.ListAutCarregamentoXNotaFiscal)
    End Sub

    Private Sub SessaoRecuperaObjEmbarquePedido()
        objEmbarquePedido = CType(Session("ssEmbarquePedido" & HID.Value), [Lib].Negocio.EmbarquePedido)
    End Sub

    Private Sub SessaoSalvaObjEmbarquePedido()
        Session("ssEmbarquePedido" & HID.Value) = objEmbarquePedido
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("TransportadorCarregamento" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtNomeProprietario.Text = itemCliente.Text
            txtCodigoProprietario.Value = itemCliente.Value
            Session.Remove("TransportadorCarregamento" & HID.Value)
            Popup.ConsultaDeTiposDeVeiculos(Me.Page, "TipoVeiculo" & HID.Value, "txtNome")

        ElseIf Session("TipoVeiculo" & HID.Value) IsNot Nothing Then
            Dim objTipo As [Lib].Negocio.TipoDeVeiculo = CType(obj, [Lib].Negocio.TipoDeVeiculo)
            txtCapacidade.Text = String.Format("{0:N4}", objTipo.Capacidade)
            txtDescricaoTipoVeiculo.Text = objTipo.Descricao
            txtCodigoTipoVeiculo.Value = objTipo.Codigo
            Session.Remove("TipoVeiculo" & HID.Value)

            Dim lstEmbarquePedido As New ListEmbarquePedido(txtCodigoProprietario.Value.Split("-"))
            If lstEmbarquePedido IsNot Nothing AndAlso lstEmbarquePedido.Count > 0 Then
                ListEmbarquePedido = lstEmbarquePedido
                SessaoSalvarListEmbarquePedido()

                'CONCATENAR PEDIDOS
                Dim pedidos As String = ""
                For Each item In lstEmbarquePedido
                    If String.IsNullOrWhiteSpace(pedidos) Then
                        pedidos &= item.CodigoPedido
                    Else
                        pedidos &= "," & item.CodigoPedido
                    End If
                Next

                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("pedido", pedidos)
                Session("ssCampo" & HID.Value) = "Carregamento"
                Session.Remove("TransportadorCarregamento" & HID.Value)
                Session("ssTipoRetorno") = "objAutCarregamento" & HID.Value
                Dim numberRows As Integer = ucConsultaPedidos.BindGridView(parameters)
                If numberRows > 1 Then
                    Popup.ConsultaDePedidos(Me.Page, "objAutCarregamento" & HID.Value)
                End If
            Else
                MsgBox(Me.Page, "Transportador sem autorização para carregamento.")
            End If
            Session.Remove("TipoVeiculo" & HID.Value)

        ElseIf Session("objPlacaAutCarregamento" & HID.Value) IsNot Nothing Then
            Dim objPlaca As [Lib].Negocio.Placa = CType(obj, [Lib].Negocio.Placa)
            Dim objTransportador As New [Lib].Negocio.Cliente(objPlaca.CodigoProprietario01, objPlaca.EndProprietario01)
            Dim itemTransportador As ListItem = Funcoes.FormatarListItemCliente(objTransportador)

            With objPlaca
                txtPlaca1.Text = .Placa01
                txtCodigoProprietario.Value = itemTransportador.Value
                txtNomeProprietario.Text = itemTransportador.Text
                txtRNTRC1.Text = .RNTRCPlaca01
                txtCodigoMotorista.Value = .CpfMotorista & "-" & .EndCpfMotorista
                txtMotorista.Text = .Motorista.Nome & " (" & .Motorista.CodigoFormatado & ")"
                txtCodigoTipoVeiculo.Value = .TipoDeVeiculoDetalhes.Codigo
                txtCapacidade.Text = String.Format("{0:N4}", .TipoDeVeiculoDetalhes.Capacidade)
                txtDescricaoTipoVeiculo.Text = .TipoDeVeiculoDetalhes.Descricao
            End With

            Dim lstEmbarquePedido As New ListEmbarquePedido(txtCodigoProprietario.Value.Split("-"))
            If lstEmbarquePedido IsNot Nothing AndAlso lstEmbarquePedido.Count > 0 Then
                ListEmbarquePedido = lstEmbarquePedido
                SessaoSalvarListEmbarquePedido()

                'CONCATENAR PEDIDOS
                Dim pedidos As String = ""
                For Each item In lstEmbarquePedido
                    If String.IsNullOrWhiteSpace(pedidos) Then
                        pedidos &= item.CodigoPedido
                    Else
                        pedidos &= "," & item.CodigoPedido
                    End If
                Next

                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("pedido", pedidos)
                Session("ssCampo" & HID.Value) = "Carregamento"
                Session.Remove("objPlacaAutCarregamento" & HID.Value)
                Session("ssTipoRetorno") = "objAutCarregamento" & HID.Value
                Dim numberRows As Integer = ucConsultaPedidos.BindGridView(parameters)
                If numberRows > 1 Then
                    Popup.ConsultaDePedidos(Me.Page, "objAutCarregamento" & HID.Value)
                End If
            Else
                MsgBox(Me.Page, "Transportador sem autorização para carregamento!")
            End If

        ElseIf Session("objAutCarregamento" & HID.Value) IsNot Nothing Then
            Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objAutCarregamento" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            SessaoRecuperaListEmbarquePedido()
            If ListEmbarquePedido IsNot Nothing AndAlso ListEmbarquePedido.Count > 0 Then
                objEmbarquePedido = ListEmbarquePedido.Where(Function(s) s.CodigoPedido = txtPedido.Text).FirstOrDefault()
                Dim lst As New ListEmbarqueXEntrega()
                ddlEntregas.Items.Clear()
                ddlEntregas.Items.Add(New ListItem("", ""))
                If objEmbarquePedido.LocaisDeEntrega.Any(Function(s) (s.Produtos.QtdeAutorizadoLE - s.Produtos.QtdeEmbarcadoLE) > 0) Then
                    For Each item In objEmbarquePedido.LocaisDeEntrega
                        If (item.Produtos.QtdeAutorizadoLE - item.Produtos.QtdeEmbarcadoLE) > 0 Then
                            lst.Add(item)
                            ddlEntregas.Items.Add(New ListItem(item.EntregaFormatado, item.CodigoClienteEntrega & "-" & item.EndClienteEntrega))
                        End If
                    Next
                Else
                    MsgBox(Me.Page, "Todas as quantidades autorizadas dos locais de entrega do pedido " & txtPedido.Text.Trim() & " já foram embarcadas!")
                End If
                objEmbarquePedido.LocaisDeEntrega = lst
                SessaoSalvaObjEmbarquePedido()
            End If
            Session.Remove("objAutCarregamento" & HID.Value)

        ElseIf Session("objNotaFiscalAut" & HID.Value) IsNot Nothing Then
            Dim lblProduto As Label = CType(grdItensCarregamento.Rows(grdItensCarregamento.SelectedIndex).FindControl("lblCodProd"), Label)

            Dim listNotas As [Lib].Negocio.ListNotasFiscais = CType(Session("objNotaFiscalAut" & HID.Value), [Lib].Negocio.ListNotasFiscais)
            Dim objCarregamentoXNota As New AutCarregamentoXNotaFiscal()
            RecuperarSessaoAutCarregamento()

            For Each objNotas As [Lib].Negocio.NotaFiscal In listNotas
                objCarregamentoXNota.Carregamento_Id = objCarregamento.Carregamento_Id
                objCarregamentoXNota.Nota_Id = objNotas.Codigo
                objCarregamentoXNota.Serie_Id = objNotas.Serie
                objCarregamentoXNota.Empresa_Id = objNotas.CodigoEmpresa
                objCarregamentoXNota.EndEmpresa_Id = objNotas.EnderecoEmpresa
                objCarregamentoXNota.Cliente_Id = objNotas.CodigoCliente
                objCarregamentoXNota.EndCliente_Id = objNotas.EnderecoCliente
                objCarregamentoXNota.EntradaSaida_Id = objNotas.EntradaSaida
                objCarregamentoXNota.CFOP_Id = objNotas.Itens(0).CFOP
                objCarregamentoXNota.Produto_Id = objNotas.Itens.Where(Function(s) s.CodigoProduto = lblProduto.Text).FirstOrDefault().CodigoProduto
                objCarregamentoXNota.Sequencia_Id = objNotas.Itens.Where(Function(s) s.CodigoProduto = lblProduto.Text).FirstOrDefault().Sequencia
                objCarregamentoXNota.IUD = "I"

                Dim objAutCarregamentoXItem As [Lib].Negocio.AutCarregamentoXItens = objCarregamento.ListCarregamentoXItens.Where(Function(s) s.CodigoProduto = lblProduto.Text).FirstOrDefault()
                If objAutCarregamentoXItem IsNot Nothing Then
                    objCarregamentoXNota.QuantidadeOrigem = objAutCarregamentoXItem.QuantidadeProgramado
                    objCarregamentoXNota.QuantidadeDestino = objAutCarregamentoXItem.QuantidadeProgramado
                End If

                If Not (listCarregamentoXNotas IsNot Nothing AndAlso listCarregamentoXNotas.Count > 0) Then
                    listCarregamentoXNotas = New [Lib].Negocio.ListAutCarregamentoXNotaFiscal()
                End If
                listCarregamentoXNotas.Add(objCarregamentoXNota)
            Next

            SessaoSalvarListCarregamentoXNotas()
            grdNotas.DataSource = listCarregamentoXNotas
            grdNotas.DataBind()
            Session.Remove("objNotaFiscalAut" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        txtPlaca1.Text = String.Empty
        txtRNTRC1.Text = String.Empty
        txtCapacidade.Text = String.Empty
        txtPedido.Text = String.Empty
        txtMotorista.Text = String.Empty
        txtCodigoMotorista.Value = String.Empty
        txtNomeProprietario.Text = String.Empty
        txtCodigoProprietario.Value = String.Empty
        txtDescricaoTipoVeiculo.Text = String.Empty
        txtCodigoTipoVeiculo.Value = String.Empty
        lblProdutoSelecionado.Text = String.Empty
        ddlEntregas.Items.Clear()

        grdRoteiro.DataSource = New List(Of Object)
        grdRoteiro.DataBind()
        grdProdutosAutorizados.DataSource = New List(Of Object)
        grdProdutosAutorizados.DataBind()
        grdItensCarregamento.DataSource = New List(Of Object)
        grdItensCarregamento.DataBind()
        grdNotas.DataSource = New List(Of Object)
        grdNotas.DataBind()

        tdNotaFiscal.Parent.Visible = False

        Session.Remove("objCarregamento" & HID.Value)
        Session.Remove("ssListEmbarque" & HID.Value)
        Session.Remove("ssEmbarquePedido" & HID.Value)
        Session.Remove("ssListCarregamentoXNotas" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaNotaFiscal.SetarHID(HID.Value)
    End Sub

    Private Function VerificaValores(ByVal qtdeInformado As Decimal, ByVal capacidade As Decimal, ByVal QtdeAutorizado As Decimal, ByVal pos As Integer) As Decimal
        If qtdeInformado = 0 Then
            MsgBox(Me.Page, "Quantidade Informada não pode ser 0")

            If capacidade < QtdeAutorizado Then
                qtdeInformado = capacidade
            Else
                qtdeInformado = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(pos).QuantidadeProgramado
            End If

        ElseIf qtdeInformado > capacidade Then
            MsgBox(Me.Page, "Quantidade Informado é maior que a capacidade do veículo")

            If capacidade < QtdeAutorizado Then
                qtdeInformado = capacidade
            Else
                qtdeInformado = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(pos).QuantidadeProgramado
            End If

        ElseIf qtdeInformado > QtdeAutorizado Then
            MsgBox(Me.Page, "Quantidade informada não pode ser maior que Quantidade Autorizada.")

            If capacidade < QtdeAutorizado Then
                qtdeInformado = capacidade
            Else
                qtdeInformado = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(pos).QuantidadeProgramado
            End If
        End If

        Return qtdeInformado
    End Function

    Private Function VerificaQtde(ByVal row As GridViewRow) As Decimal
        Try
            Dim Capacidade As Decimal = txtCapacidade.Text
            Dim QtdeAutorizado As Decimal = CType(row.FindControl("lblQtdeAut"), Label).Text

            If Capacidade < QtdeAutorizado Then
                Return Capacidade
            Else
                Return QtdeAutorizado
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                Limpar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaPlaca_Click(sender As Object, e As EventArgs) Handles btnBuscaPlaca.Click
        Try
            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlacaAutCarregamento" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
            ucConsultaPlacas.Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaProprietario_Click(sender As Object, e As EventArgs) Handles btnBuscaProprietario.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtPlaca1.Text) Then
                MsgBox(Me.Page, "Placa já foi selecionada, limpe os campos para nova consulta.")
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
                Popup.ConsultaDeClientes(Me.Page, "TransportadorCarregamento" & HID.Value, "txtNome")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEntregas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEntregas.SelectedIndexChanged
        Try
            If ddlEntregas.SelectedIndex > 0 Then

                SessaoRecuperaObjEmbarquePedido()

                grdRoteiro.DataSource = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).Roteiros
                grdRoteiro.DataBind()

                grdProdutosAutorizados.DataSource = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).Produtos
                grdProdutosAutorizados.DataBind()

                'grdItensCarregamento.DataSource = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos(0).ListCarregamentoXItens
                'grdItensCarregamento.DataBind()

                SessaoSalvaObjEmbarquePedido()
            Else
                grdRoteiro.DataSource = New List(Of Object)
                grdRoteiro.DataBind()
                grdProdutosAutorizados.DataSource = New List(Of Object)
                grdProdutosAutorizados.DataBind()
                grdItensCarregamento.DataSource = New List(Of Object)
                grdItensCarregamento.DataBind()
                grdNotas.DataSource = New List(Of Object)
                grdNotas.DataBind()
                tdNotaFiscal.Parent.Visible = False
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

    Protected Sub lnkAdicionaAutorizacao_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SessaoRecuperaObjEmbarquePedido()
            RecuperarSessaoAutCarregamento()

            Dim objEmbarqueEntrega As New EmbarqueXEntrega(objEmbarquePedido)
            Dim row As GridViewRow = CType(CType(sender, LinkButton).NamingContainer, GridViewRow)

            If objCarregamento Is Nothing Then
                objCarregamento = New AutCarregamento(objEmbarqueEntrega)
                objCarregamento.IUD = "I"
                objCarregamento.Carregamento_Id = IIf(objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Max(Function(s) s.Carregamento_Id) Is Nothing, 1, objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Max(Function(s) s.Carregamento_Id) + 1)
                objCarregamento.Placa = txtPlaca1.Text
                objCarregamento.PercAdiantamento = 0
                objCarregamento.Situacao = 1
                objCarregamento.MotivoCancelamento = String.Empty
                objCarregamento.NrCarregamentoTerceiro = 1
                objCarregamento.UsuarioInclusao = UsuarioServidor.NomeUsuario
                objCarregamento.UsuarioInclusaoData = DateTime.Now
                objCarregamento.UsuarioCancelamento = String.Empty
                objCarregamento.UsuarioCancelamentoData = New Nullable(Of DateTime)
                objCarregamento.Movimento = DateTime.Now
                objCarregamento.CodigoTransportador = txtCodigoProprietario.Value.Split("-")(0)
                objCarregamento.EndTransportador = txtCodigoProprietario.Value.Split("-")(1)
                objCarregamento.CodigoMotorista = txtCodigoMotorista.Value.Split("-")(0)
                objCarregamento.EndMotorista = txtCodigoMotorista.Value.Split("-")(1)
                objCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega = ddlEntregas.SelectedValue.Split("-")(0)
                objCarregamento.ParentEmbarqueEntrega.EndClienteEntrega = ddlEntregas.SelectedValue.Split("-")(1)
                objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Add(objCarregamento)
            End If

            If objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos().Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).SelectMany(Function(s) s.ListCarregamentoXItens).Any(Function(s) s.CodigoProduto = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).Produtos(row.RowIndex).CodigoProduto) Then
                MsgBox(Me.Page, "Produto já inserido neste carregamento.")
                Exit Sub
            End If

            Dim nrCotacao As Integer
            Dim listEmbPrecos As New ListEmbarquePrecoFrete(objCarregamento)

            For Each preco As EmbarquePrecoFrete In listEmbPrecos
                Dim qtdeCarregado As Decimal = New ListAutCarregamentoXItens(objCarregamento, preco.NrCotacao).Sum(Function(s) s.QuantidadeProgramado)
                Dim result As Integer = preco.Quota - qtdeCarregado
                If result > 0 Then
                    nrCotacao = preco.NrCotacao
                    Exit For
                End If
            Next

            Dim objCarregamentoXItens As New AutCarregamentoXItens(objCarregamento)
            Dim qtdeCarregar As Decimal = VerificaQtde(row)

            objCarregamentoXItens.IUD = "I"
            objCarregamentoXItens.CodigoProduto = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).Produtos(row.RowIndex).CodigoProduto
            objCarregamentoXItens.NrCotacao_Id = nrCotacao
            objCarregamentoXItens.QuantidadeProgramado = qtdeCarregar
            objCarregamentoXItens.PesoProgramado = qtdeCarregar
            objCarregamentoXItens.VolumesProgramado = qtdeCarregar
            objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens.Add(objCarregamentoXItens)

            grdItensCarregamento.DataSource = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).SelectMany(Function(s) s.ListCarregamentoXItens)
            grdItensCarregamento.DataBind()

            SalvarSessaoAutCarregamento()
            SessaoSalvaObjEmbarquePedido()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            SessaoRecuperaObjEmbarquePedido()
            SessaoRecuperaListCarregamentoXNotas()

            If ddlEntregas.SelectedIndex <> -1 AndAlso objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Count > 0 AndAlso objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos(0).ListCarregamentoXItens.Count > 0 Then
                Dim Sqls As New ArrayList

                For Each objCarregto As [Lib].Negocio.AutCarregamento In objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos
                    If String.IsNullOrWhiteSpace(objCarregto.IUD) Then
                        objCarregto.IUD = "U"
                        'ElseIf objCarregto.IUD = "I" Then
                        '    objCarregto.Carregamento_Id = IIf(objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Max(Function(s) s.Carregamento_Id) Is Nothing, 1, objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Max(Function(s) s.Carregamento_Id) + 1)
                    End If
                    If objCarregto.IUD <> "U" Then
                        objCarregto.SalvarSql(Sqls)
                    End If
                Next

                If listCarregamentoXNotas IsNot Nothing AndAlso listCarregamentoXNotas.Count > 0 Then
                    For Each objCarregamentoXNotas As AutCarregamentoXNotaFiscal In listCarregamentoXNotas
                        objCarregamentoXNotas.SalvarSql(Sqls)
                    Next
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                Else
                    MsgBox(Me.Page, "Carregamento liberado.")
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Não foi adicionado itens ao carregamento.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtPec_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim row As GridViewRow = CType(CType(sender, TextBox).NamingContainer, GridViewRow)
            Dim txtPercAdiantamento As TextBox = CType(sender, TextBox)
            SessaoRecuperaObjEmbarquePedido()
            RecuperarSessaoAutCarregamento()

            If Not String.IsNullOrWhiteSpace(txtPercAdiantamento.Text) Then
                objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().PercAdiantamento = txtPercAdiantamento.Text
            End If

            SessaoSalvaObjEmbarquePedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtQtdeAutorizado_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim row As GridViewRow = CType(CType(sender, TextBox).NamingContainer, GridViewRow)
            Dim txtQtdeInformado As TextBox = CType(sender, TextBox)
            SessaoRecuperaObjEmbarquePedido()
            RecuperarSessaoAutCarregamento()
            Dim QtdeAutorizado As Decimal

            For Each rowProduto As GridViewRow In grdProdutosAutorizados.Rows
                If CType(rowProduto.FindControl("lblCodProd"), Label).Text = objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(0).CodigoProduto Then
                    QtdeAutorizado = CType(rowProduto.FindControl("lblQtdeAut"), Label).Text
                End If
            Next

            txtQtdeInformado.Text = VerificaValores(txtQtdeInformado.Text, txtCapacidade.Text, QtdeAutorizado, row.RowIndex)

            If Not String.IsNullOrWhiteSpace(txtQtdeInformado.Text) AndAlso txtQtdeInformado.Text > 0 Then
                objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(row.RowIndex).QuantidadeProgramado = txtQtdeInformado.Text
                objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(row.RowIndex).PesoProgramado = txtQtdeInformado.Text
                objEmbarquePedido.LocaisDeEntrega(ddlEntregas.SelectedIndex - 1).ListCarregamentos.Where(Function(s) s.Carregamento_Id = objCarregamento.Carregamento_Id).FirstOrDefault().ListCarregamentoXItens(row.RowIndex).VolumesProgramado = txtQtdeInformado.Text
            End If

            SessaoSalvaObjEmbarquePedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNotaFiscal_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNotaFiscal.Click
        Try
            If grdItensCarregamento.SelectedIndex < 0 Then
                MsgBox(Me.Page, "Selecione o produto.")
                Exit Sub
            End If

            Dim lblProd As Label = CType(grdItensCarregamento.Rows(grdItensCarregamento.SelectedIndex).FindControl("lblCodProd"), Label)

            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("pedido", txtPedido.Text.Trim())
            parameters.Add("produto", lblProd.Text.Trim())
            ucConsultaNotaFiscal.Limpar()
            ucConsultaNotaFiscal.BindGridView(parameters)
            Popup.ConsultaNotaFiscal(Me.Page, "objNotaFiscalAut" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirItemCarregamento_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperarSessaoAutCarregamento()
            SessaoRecuperaListCarregamentoXNotas()

            Dim row As GridViewRow = CType(sender.NamingContainer, GridViewRow)

            If listCarregamentoXNotas IsNot Nothing AndAlso listCarregamentoXNotas.Count > 0 Then

                For i = 0 To listCarregamentoXNotas.Count - 1
                    If listCarregamentoXNotas(i).Produto_Id = objCarregamento.ListCarregamentoXItens(row.RowIndex).CodigoProduto Then
                        listCarregamentoXNotas.RemoveAt(i)
                    End If
                Next

            End If

            objCarregamento.ListCarregamentoXItens.RemoveAt(row.RowIndex)

            grdItensCarregamento.DataSource = objCarregamento.ListCarregamentoXItens
            grdItensCarregamento.DataBind()

            grdNotas.DataSource = listCarregamentoXNotas
            grdNotas.DataBind()

            If listCarregamentoXNotas.Count = 0 Then
                tdNotaFiscal.Parent.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirItemCarregamentoXNota_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaListCarregamentoXNotas()

            Dim row As GridViewRow = CType(sender.NamingContainer, GridViewRow)
            listCarregamentoXNotas.RemoveAt(row.RowIndex)

            grdNotas.DataSource = listCarregamentoXNotas
            grdNotas.DataBind()

            If listCarregamentoXNotas.Count = 0 Then
                tdNotaFiscal.Parent.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdItensCarregamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles grdItensCarregamento.SelectedIndexChanged
        Try
            SessaoRecuperaListCarregamentoXNotas()
            RecuperarSessaoAutCarregamento()

            Dim lblProd As Label = CType(grdItensCarregamento.Rows(grdItensCarregamento.SelectedIndex).FindControl("lblCodProd"), Label)
            Dim lblDesc As Label = CType(grdItensCarregamento.Rows(grdItensCarregamento.SelectedIndex).FindControl("lblDescProd"), Label)

            lblProdutoSelecionado.Text = lblProd.Text & "-" & lblDesc.Text
            tdNotaFiscal.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AutorizacaoDeCarregamento")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class