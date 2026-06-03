Imports NGS.Lib.Negocio
Imports System.Threading
Imports System.IO

Public Class AlterarProdutoNotaGerais
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

#Region "Property"

    Public Property index() As Integer
        Get
            Return Session("index")
        End Get
        Set(ByVal value As Integer)
            Session("index") = value
        End Set
    End Property

#End Region

#Region "Events"


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If IsConnect AndAlso Not IsPostBack Then
                If Funcoes.VerificaPermissao("AlterarProdutoNFG", "ACESSAR") Then
                    CargaUnidadeDeNegocio()
                    limpar()
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    Dim fm As New FilesManager()
                    If Not fm.IsConnect() Then
                        MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
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
            If Funcoes.VerificaPermissao("AlterarProdutoNFG", "LEITURA") Then

                RecuperaNotaFiscal()
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

                If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    MsgBox(Me.Page, "É necessário selecionar a empresa!")
                    Exit Sub
                End If

                objNotaFiscal.CodigoEmpresa = Empresa(0)
                objNotaFiscal.EnderecoEmpresa = Empresa(1)
                objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                objNotaFiscal.CarregandoNota = True
                'objNotaFiscal.NossaEmissao = True
                objNotaFiscal.CodigoSituacao = 1
                objNotaFiscal.CodigoTipoDeDocumento = 1

                If txtCodigoCliente.Value.ToString.Length > 0 Then
                    Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                    objNotaFiscal.CodigoCliente = Cliente(0)
                    objNotaFiscal.EnderecoCliente = Cliente(1)
                End If

                'AlterarNotasEmDemanda(sender, e)
                'Exit Sub

                If txtES.Text.Length > 0 Then objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)

                If txtNota.Text.Length > 0 Then
                    objNotaFiscal.Codigo = txtNota.Text
                Else
                    objNotaFiscal.Codigo = 0
                End If

                If txtSerie.Text.Length > 0 Then
                    objNotaFiscal.Serie = txtSerie.Text
                Else
                    objNotaFiscal.Serie = ""
                End If

                Session("ssCampo" & HID.Value) = "AlterarNotaGeral"

                ucConsultaPedidosXNotas.BindGridView()
                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNFG" & HID.Value.ToString)
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AlterarNotasEmDemanda(sender As Object, e As EventArgs)


        'New NotaParaAlteracao("181824", "E", "301010003", "301012683"),
        'New NotaParaAlteracao("181835", "E", "301010012", "301010005"),
        Dim list As New List(Of String)
        Dim notas As New List(Of NotaParaAlteracao) From {
            New NotaParaAlteracao("181975", "E", "301010032", "301012808"),
            New NotaParaAlteracao("181921", "E", "301010035", "301010025"),
            New NotaParaAlteracao("181973", "E", "301010038", "301010927"),
            New NotaParaAlteracao("181980", "E", "301010052", "301010936"),
            New NotaParaAlteracao("181983", "E", "301010054", "301010938"),
            New NotaParaAlteracao("181985", "E", "301010056", "301012892"),
            New NotaParaAlteracao("181703", "E", "301010065", "301012083"),
            New NotaParaAlteracao("181773", "E", "301010184", "301010194"),
            New NotaParaAlteracao("181548", "E", "301010192", "301010177"),
            New NotaParaAlteracao("183600", "E", "301010298", "301010300"),
            New NotaParaAlteracao("183641", "E", "301010387", "301012710"),
            New NotaParaAlteracao("181395", "E", "301010397", "301013117"),
            New NotaParaAlteracao("183675", "E", "301010451", "301012965"),
            New NotaParaAlteracao("183687", "E", "301010460", "301012942"),
            New NotaParaAlteracao("183685", "E", "301010463", "301012875"),
            New NotaParaAlteracao("183691", "E", "301010465", "301012838"),
            New NotaParaAlteracao("181008", "E", "301010481", "301012979"),
            New NotaParaAlteracao("183714", "E", "301010488", "301012752"),
            New NotaParaAlteracao("183716", "E", "301010490", "301010504"),
            New NotaParaAlteracao("247", "E", "301010492", "301012985"),
            New NotaParaAlteracao("183706", "E", "301010492", "301010480"),
            New NotaParaAlteracao("183720", "E", "301010492", "301012925"),
            New NotaParaAlteracao("183721", "E", "301010492", "301012985"),
            New NotaParaAlteracao("183728", "E", "301010498", "301012690"),
            New NotaParaAlteracao("181967", "E", "301010500", "301010739"),
            New NotaParaAlteracao("181819", "E", "301010511", "301010514"),
            New NotaParaAlteracao("181817", "E", "301010514", "301010512"),
            New NotaParaAlteracao("181311", "E", "301010594", "301010126"),
            New NotaParaAlteracao("181913", "E", "301010739", "301012899"),
            New NotaParaAlteracao("181919", "E", "301010759", "301012706"),
            New NotaParaAlteracao("181891", "E", "301010829", "301012907"),
            New NotaParaAlteracao("181048", "E", "301010860", "301010301"),
            New NotaParaAlteracao("180943", "E", "301010903", "301012643"),
            New NotaParaAlteracao("181091", "E", "301011075", "301010493"),
            New NotaParaAlteracao("181900", "E", "301011084", "301010681"),
            New NotaParaAlteracao("181467", "E", "301011098", "301012840"),
            New NotaParaAlteracao("181905", "E", "301011098", "301012661"),
            New NotaParaAlteracao("182033", "E", "301011130", "301013116"),
            New NotaParaAlteracao("181539", "E", "301011239", "301012237"),
            New NotaParaAlteracao("183016", "E", "301011248", "301012884"),
            New NotaParaAlteracao("181044", "E", "301011880", "301013116"),
            New NotaParaAlteracao("183496", "E", "301012001", "301012492"),
            New NotaParaAlteracao("181487", "E", "301012033", "301012034"),
            New NotaParaAlteracao("183176", "E", "301012118", "301012121"),
            New NotaParaAlteracao("183174", "E", "301012120", "301012122"),
            New NotaParaAlteracao("183172", "E", "301012122", "301012118"),
            New NotaParaAlteracao("178", "E", "301012155", "301012203"),
            New NotaParaAlteracao("181703", "E", "301012261", "301012923"),
            New NotaParaAlteracao("181389", "E", "301012274", "301012327"),
            New NotaParaAlteracao("183327", "E", "301012279", "301012596"),
            New NotaParaAlteracao("183", "E", "301012289", "301012494"),
            New NotaParaAlteracao("183285", "E", "301012294", "301013117"),
            New NotaParaAlteracao("183406", "E", "301012351", "301012687"),
            New NotaParaAlteracao("181419", "E", "301012358", "301013118"),
            New NotaParaAlteracao("183416", "E", "301012358", "301013118")
        }

        For Each nota In notas

            objNotaFiscal.Codigo = nota.Nota
            objNotaFiscal.EntradaSaida = IIf(nota.EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)

            Dim sql As String
            Dim ds As DataSet

            sql = ucConsultaPedidosXNotas.SQLAlterarNotasGerais()
            ds = Banco.ConsultaDataSet(sql, "Notas")

            If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count = 1 Then

                Dim objNFConsulta As New [Lib].Negocio.NotaFiscal()

                objNFConsulta.CodigoEmpresa = ds.Tables(0).Rows(0).Item("Empresa_Id")
                objNFConsulta.EnderecoEmpresa = ds.Tables(0).Rows(0).Item("EndEmpresa_Id")
                objNFConsulta.CodigoCliente = ds.Tables(0).Rows(0).Item("Cliente")
                objNFConsulta.EnderecoCliente = ds.Tables(0).Rows(0).Item("EndCliente")
                objNFConsulta.EntradaSaida = IIf(ds.Tables(0).Rows(0).Item("ES") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                objNFConsulta.Serie = ds.Tables(0).Rows(0).Item("Serie")
                objNFConsulta.Codigo = ds.Tables(0).Rows(0).Item("Nota")
                objNFConsulta.ChaveNFE = ""
                objNFConsulta.Eletronica = False

                Session("objNFConsultaNFG" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)

                CarregarNotaFiscal()

                Dim objProduto As New [Lib].Negocio.Produto(nota.ProdutoCorrigido)
                Session("objProduto" & HID.Value) = objProduto
                Carregar(objProduto)

                lnkAtualizar_Click(sender, e)

            ElseIf ds.Tables(0).Rows.Count > 1 Then

                list.Add(nota.Nota)

            End If

        Next

    End Sub

    Protected Sub chkProduto_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim chk As CheckBox = CType(sender, CheckBox)

            If chk.Checked Then
                Dim row As GridViewRow = CType(chk.NamingContainer, GridViewRow)
                index = row.RowIndex

                Session("Where" & HID.Value) = " Situacao = 1 "
                ucConsultaProduto.Limpar()
                ucConsultaProduto.BuscarProduto()

                Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)

                Popup.ConsultaDeProduto(Me.Page, "objProduto" & HID.Value, txtNome.ClientID, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkObservacaoProd_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim chk As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chk.NamingContainer, GridViewRow)
            Dim txtObservacaoProd As TextBox = CType(row.FindControl("txtObservacaoProd"), TextBox)

            txtObservacaoProd.Enabled = chk.Checked
            txtObservacaoProd.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click

        Try
            If Funcoes.VerificaPermissao("AlterarProdutoNFG", "ALTERAR") Then
                RecuperaNotaFiscal()

                If ValidaCampos() Then

                    Dim Sqls As New ArrayList
                    Dim msgNFE As String = String.Empty
                    Dim Sql As String = String.Empty

                    For Each row In grdNFGProdutos.Rows
                        If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                            Dim nI As Integer = CType(CType(row.FindControl("chkProduto"), CheckBox).NamingContainer, GridViewRow).RowIndex
                            If objNotaFiscal.Itens(nI).NotasDevolucao IsNot Nothing AndAlso objNotaFiscal.Itens(nI).NotasDevolucao.Count > 0 Then
                                MsgBox(Me.Page, "Produto vinculado a Devolução não pode ser alterado/corrigido.", eTitulo.Info)
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If

                            Sql = "SELECT NotaReferencial_Id " & vbCrLf &
                                    "  FROM NotaFiscalReferencial" & vbCrLf &
                                    " WHERE EmpresaReferencial_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                    "   AND EndEmpresaReferencial_Id   =" & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                    "   AND ClienteReferencial_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                    "   AND EndClienteReferencial_Id   =" & objNotaFiscal.EnderecoCliente & vbCrLf &
                                    "   AND EntradaSaidaReferencial_Id ='" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                                    "   AND SerieReferencial_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                                    "   AND NotaReferencial_Id         =" & objNotaFiscal.Codigo & vbCrLf &
                                    "   AND ProdutoReferencial_Id      ='" & objNotaFiscal.Itens(nI).CodigoProduto & "'" & vbCrLf &
                                    "   AND CFOPReferencial_Id         =" & objNotaFiscal.Itens(nI).CFOP & vbCrLf &
                                    "   AND SequenciaReferencial_Id    =" & objNotaFiscal.Itens(nI).Sequencia

                            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotaFiscalReferencial1")

                            If ds IsNot Nothing AndAlso ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                                MsgBox(Me.Page, "Produto " & objNotaFiscal.Itens(nI).CodigoProduto & " com Nota Fiscal Referenciada não pode ser alterada/corrigida.")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If

                            Sql = "SELECT NotaReferencial_Id " & vbCrLf &
                                    "  FROM NotaFiscalReferencial" & vbCrLf &
                                    " WHERE Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                    "   AND EndEmpresa_Id   =" & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                    "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                    "   AND EndCliente_Id   =" & objNotaFiscal.EnderecoCliente & vbCrLf &
                                    "   AND EntradaSaida_Id ='" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                                    "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                                    "   AND Nota_Id         =" & objNotaFiscal.Codigo & vbCrLf &
                                    "   AND Produto_Id      ='" & objNotaFiscal.Itens(nI).CodigoProduto & "'" & vbCrLf &
                                    "   AND CFOP_Id         =" & objNotaFiscal.Itens(nI).CFOP & vbCrLf &
                                    "   AND Sequencia_Id    =" & objNotaFiscal.Itens(nI).Sequencia

                            Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "NotaFiscalReferencial2")

                            If ds2 IsNot Nothing AndAlso ds2.Tables(0) IsNot Nothing AndAlso ds2.Tables(0).Rows.Count > 0 Then
                                MsgBox(Me.Page, "Produto " & objNotaFiscal.Itens(nI).CodigoProduto & " com Nota Fiscal Referenciada não pode ser alterada/corrigida.")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If

                            If objNotaFiscal.Itens(nI).Lotes IsNot Nothing AndAlso objNotaFiscal.Itens(nI).Lotes.Count > 0 Then
                                For Each nXlt In objNotaFiscal.Itens(nI).Lotes
                                    Sql = "Select Ordem_Id FROM OrdemDeProducaoXConsumoXLote" & vbCrLf &
                                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                            "  and Produto_Id      ='" & objNotaFiscal.Itens(nI).CodigoProduto & "'" & vbCrLf &
                                            "  and Lote_Id         ='" & nXlt.Lote & "'"

                                    Dim dsL As DataSet = Banco.ConsultaDataSet(Sql, "OrdemDeProducaoXConsumoXLote")

                                    If dsL IsNot Nothing AndAlso dsL.Tables(0) IsNot Nothing AndAlso dsL.Tables(0).Rows.Count > 0 Then
                                        MsgBox(Me.Page, "Produto " & objNotaFiscal.Itens(nI).CodigoProduto & " do Lote " & nXlt.Lote & " está lançado na Ordem " & dsL.Tables(0).Rows(0)(0).ToString & " e não pode ser alterado/corrigido.")
                                        lnkAtualizar.Parent.Visible = False
                                        Exit Sub
                                    End If
                                Next
                            End If
                        End If
                    Next

                    Sql = " Update NotasXEmbalagens set " & vbCrLf &
                             "      Quantidade            = " & Str(objNotaFiscal.Quantidade) & "," & vbCrLf &
                             "      Especie               = '" & objNotaFiscal.Especie & "'," & vbCrLf &
                             "      Marca                 = '" & objNotaFiscal.Marca & "'," & vbCrLf &
                             "      Numero                = '" & objNotaFiscal.Numero & "'," & vbCrLf &
                             "      PesoBruto             = " & Str(objNotaFiscal.PesoBruto) & "," & vbCrLf &
                             "      PesoLiquido           = " & Str(objNotaFiscal.PesoLiquido) & vbCrLf &
                             "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                             "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                             "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                             "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                             "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                             "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                             "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                    Sqls.Add(Sql)

                    Dim temProdutoPedido As Boolean = False
                    For Each item In objNotaFiscal.Pedido.Itens
                        If item.CodigoProduto = Trim(CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text) Then
                            temProdutoPedido = True
                            Exit For
                        End If
                    Next

                    For Each row In grdNFGProdutos.Rows
                        If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                            index = CType(CType(row.FindControl("chkProduto"), CheckBox).NamingContainer, GridViewRow).RowIndex

                            'Incluir Produto no Pedido e Nota Fiscal
                            If Not temProdutoPedido Then

                                Sql = "SELECT      FatorConversao" & vbCrLf &
                                      "FROM        PedidoXItem" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  AND EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  AND Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  AND Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'; "

                                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "PedidoXItem")

                                If ds IsNot Nothing AndAlso ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                                    For Each rowPedido As DataRow In ds.Tables(0).Rows
                                        If rowPedido("FatorConversao") <> CType(grdNFGProdutos.Rows(index).FindControl("txtFator"), TextBox).Text Then
                                            MsgBox(Me.Page, "Produto " & objNotaFiscal.Itens(index).CodigoProduto & " tem um fator de conversão diferente do produto a ser alterado! Entre em contato com o suporte.")
                                            lnkAtualizar.Parent.Visible = False
                                            Exit Sub
                                        End If
                                    Next

                                End If

                                Dim dFator As Decimal = CType(grdNFGProdutos.Rows(index).FindControl("txtFator"), TextBox).Text

                                Sql = "insert into PedidoXItem (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Classificacao, Retencao, OperacaoXEstado, UnidadeComercializacao, FatorConversao)" & vbCrLf &
                                      "Select      Empresa_Id, EndEmpresa_Id, Pedido_Id, '" & CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', Classificacao, Retencao, " & CType(grdNFGProdutos.Rows(index).FindControl("txtOperacaoEstado"), TextBox).Text & ", '" & CType(grdNFGProdutos.Rows(index).FindControl("txtUnidade"), TextBox).Text & "', " & Str(dFator) & "" & vbCrLf &
                                      "FROM        PedidoXItem" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "insert into  PedidoXItemXLancamento (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, PedidoItem_Id, TipoDeLancamento, Movimento, Quantidade, UnitarioOficial, UnitarioMoeda, TotalOficial, " & vbCrLf &
                                      "                      TotalMoeda, DataEntrega, UnitarioOficialCompra, UnitarioMoedaCompra, UsuarioLiberacao, UsuarioLiberacaoData, QuantidadeComercializacao)" & vbCrLf &
                                      "SELECT     Empresa_Id, EndEmpresa_Id, Pedido_Id, '" & CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', PedidoItem_Id, TipoDeLancamento, Movimento, Quantidade, UnitarioOficial, UnitarioMoeda, TotalOficial, " & vbCrLf &
                                      "                      TotalMoeda, DataEntrega, UnitarioOficialCompra, UnitarioMoedaCompra, UsuarioLiberacao, UsuarioLiberacaoData, QuantidadeComercializacao" & vbCrLf &
                                      "FROM         PedidoXItemXLancamento" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "insert into PedidosXEncargos (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Encargo_Id, BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda)" & vbCrLf &
                                      "SELECT      Empresa_Id, EndEmpresa_Id, Pedido_Id, '" & CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', Encargo_Id, BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda" & vbCrLf &
                                      "FROM        PedidosXEncargos" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)
                            End If

                            Sql = "insert into NotasFiscaisXItens (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Pedido, Deposito, EndDeposito, DepositoTerceiro, " & vbCrLf &
                            "                      EndDepositoTerceiro, PesoFiscal, QuantidadeFisica, QuantidadeFiscal, Unitario, Valor, PesoQuantidade, Operacao, SubOperacao, Lote, ObservacoesDoProduto, " & vbCrLf &
                            "                      Classificacao, Sequencia_Id, Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeDeEmbalagem, Rateado, NumeroDeclaracaoImportacao, DataRegistroDI, " & vbCrLf &
                            "                      LocalDesembarqueImportacao, EstadoDesembarqueImportacao, DataDesembarqueImportacao, Fabricante, EndFabricante, NumeroPecas, Fixacao, ValorMoeda, " & vbCrLf &
                            "                      ValorLiquido, ValorLiquidoMoeda, OperacaoXEstado)" & vbCrLf &
                            "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, '" & CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', CFOP_Id, Pedido, Deposito, EndDeposito, DepositoTerceiro, " & vbCrLf &
                            "                      EndDepositoTerceiro, PesoFiscal, QuantidadeFisica, QuantidadeFiscal, Unitario, Valor, PesoQuantidade, Operacao, SubOperacao, Lote," & vbCrLf

                            If CType(row.FindControl("chkObservacaoProd"), CheckBox).Checked Then
                                Sql &= "'" & Funcoes.EliminarCaracteresEspeciaisNF(CType(grdNFGProdutos.Rows(index).FindControl("txtObservacaoProd"), TextBox).Text) & "'," & vbCrLf
                            Else
                                Sql &= " ObservacoesDoProduto, " & vbCrLf
                            End If

                            Sql &= "                      Classificacao, Sequencia_Id, Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeDeEmbalagem, Rateado, NumeroDeclaracaoImportacao, DataRegistroDI, " & vbCrLf &
                            "                      LocalDesembarqueImportacao, EstadoDesembarqueImportacao, DataDesembarqueImportacao, Fabricante, EndFabricante, NumeroPecas, Fixacao, ValorMoeda, " & vbCrLf &
                            "                      ValorLiquido, ValorLiquidoMoeda, " & CType(grdNFGProdutos.Rows(index).FindControl("txtOperacaoEstado"), TextBox).Text & vbCrLf &
                            "FROM         NotasFiscaisXItens" & vbCrLf &
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            Sql = "insert into NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Encargo_Id, Base, Percentual, Valor, " & vbCrLf &
                            "                      CentroDeCusto, Sequencia_id, PercentualExibicao)" & vbCrLf &
                            "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, '" & CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', CFOP_Id, Encargo_Id, Base, Percentual, Valor, " & vbCrLf &
                            "                      CentroDeCusto, Sequencia_id, PercentualExibicao" & vbCrLf &
                            "FROM         NotasFiscaisXEncargos" & vbCrLf &
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            If objNotaFiscal.Itens(index).Lotes IsNot Nothing AndAlso objNotaFiscal.Itens(index).Lotes.Count > 0 Then
                                For Each nXl In objNotaFiscal.Itens(index).Lotes

                                    Sql = "insert into NotaFiscalXLote (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Lote_Id, Quantidade, Validade, QuantidadeDeConsumo)" & vbCrLf &
                                    "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, '" & CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', " & vbCrLf &
                                    "       CFOP_Id, Sequencia_Id, Lote_Id, Quantidade, Validade, QuantidadeDeConsumo" & vbCrLf &
                                    "FROM         NotaFiscalXLote" & vbCrLf &
                                    "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                    "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                    "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                    "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                                    "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                    "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                                    "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                                    "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                                    "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                                    "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'" & vbCrLf &
                                    "  and Lote_Id         ='" & nXl.Lote & "'"
                                    Sqls.Add(Sql)

                                Next
                            End If


                            'Remover Produto da Nota Fiscal e do Pedido
                            Sql = "DELETE FROM NotasFiscaisXEncargos" & vbCrLf &
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            If objNotaFiscal.Itens(index).Lotes IsNot Nothing AndAlso objNotaFiscal.Itens(index).Lotes.Count > 0 Then
                                For Each nXl In objNotaFiscal.Itens(index).Lotes

                                    Sql = "DELETE FROM NotaFiscalXLote" & vbCrLf &
                                    "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                    "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                    "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                    "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                                    "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                    "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                                    "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                                    "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                                    "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                                    "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'" & vbCrLf &
                                    "  and Lote_Id         ='" & nXl.Lote & "'"
                                    Sqls.Add(Sql)

                                Next
                            End If

                            Sql = "DELETE FROM NotasFiscaisXItens" & vbCrLf &
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            If Not temProdutoPedido Then
                                Sql = "DELETE FROM PedidosXEncargos" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "DELETE FROM PedidoXItemXLancamento" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "DELETE FROM        PedidoXItem" & vbCrLf &
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf &
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)
                            End If
                        ElseIf CType(row.FindControl("chkObservacaoProd"), CheckBox).Checked Then
                            index = CType(CType(row.FindControl("chkObservacaoProd"), CheckBox).NamingContainer, GridViewRow).RowIndex
                            Sql = " Update NotasFiscaisXItens set " & vbCrLf &
                                 "      ObservacoesDoProduto = '" & Funcoes.EliminarCaracteresEspeciaisNF(CType(grdNFGProdutos.Rows(index).FindControl("txtObservacaoProd"), TextBox).Text) & "'" & vbCrLf &
                                 "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                 "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                 "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                 "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                                 "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                 "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                                 "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                                 "    and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf &
                                 "    and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf &
                                 "    and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)
                        End If
                    Next

                    Dim obs As String = objNotaFiscal.ObservacoesControleInterno

                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Alteração do Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Alteração do Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    End If

                    Sql = " Update NotasFiscais set " & vbCrLf &
                         "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                         "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
                         "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
                         "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                         "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                         "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                         "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                         "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                         "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                         "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                    Sqls.Add(Sql)

                    If Banco.GravaBanco(Sqls) Then
                        SalvaNotaFiscal()
                    Else
                        Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                        MsgBox(Me.Page, Observacao)
                        Exit Sub
                    End If

                    Thread.Sleep(5000)

                    Sqls.Clear()

                    RecuperaNotaFiscal()

                    For Each row In grdNFGProdutos.Rows
                        If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                            objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)

                            objNotaFiscal.NotaFiscalOriginal.Razao.ExcluiContabilizacaoNotaSQL(Sqls)

                            objNotaFiscal.NotaFiscalOriginal.Razao.ContabilizarNotaSql(Sqls)

                            Exit For
                        End If
                    Next

                    If Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, "Produto alterado com sucesso. Verifique o Extrato do Pedido se a alteração ficou de acordo e o Espelho da Nota Fiscal se os Valores e Contabilização estão corretos.", eTitulo.Sucess)
                        Extrato.Emitir(Me.Page, False, objNotaFiscal.Pedido.CodigoEmpresa, objNotaFiscal.Pedido.EnderecoEmpresa, "T",
                                                                      objNotaFiscal.Pedido.Codigo)

                        Dim espelho As New NotaFiscalEspelho()
                        espelho.ExibirEspelho(Me.Page, objNotaFiscal.NotaFiscalOriginal)

                        limpar()
                    Else
                        Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                        MsgBox(Me.Page, Observacao)
                    End If

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            Dim objProduto As [Lib].Negocio.Produto = Session("objProduto" & HID.Value)
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objCliente" & HID.Value)
            ElseIf Session("objProduto" & HID.Value) IsNot Nothing Then
                RecuperaNotaFiscal()

                'buscar operacoes x encargos desse produto (objProduto.codigo)
                Dim Parametros As New OperacaoXEstado
                Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
                Parametros.CodigoGrupoProduto = objProduto.CodigoGrupo
                Parametros.CodigoProduto = objProduto.Codigo
                Parametros.CodigoOperacao = objNotaFiscal.Itens(index).CodigoOperacao
                Parametros.CodigoSubOperacao = objNotaFiscal.Itens(index).CodigoSubOperacao
                Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado
                Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado

                If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao

                Dim newObjOE As New OperacaoXEstado(Parametros)

                If newObjOE.Encargos.Count > 0 Then
                    For Each enc In objNotaFiscal.Itens(index).Encargos
                        If newObjOE.Encargos.Any(Function(s) s.CodigoEncargo = enc.Codigo And s.OperacaoEstado.CodigoFiscal = enc.ItemNotaFiscal.CFOP) Then
                            If newObjOE.CodigoSTICMS = enc.SituacaoTributaria AndAlso
                                newObjOE.CodigoSTIPI = enc.SituacaoTributariaIPI AndAlso
                                newObjOE.CodigoSTPISCOFINS = enc.SituacaoTributariaPISCOFINS Then
                            ElseIf objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                                CType(grdNFGProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = False
                                MsgBox(Me.Page, "Encargo " & enc.Codigo & " com Situação (Tributária, IPI ou PIS/COFINS) diferente do Produto selecionado.")
                                Exit Sub
                            End If
                        Else
                            CType(grdNFGProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = False
                            MsgBox(Me.Page, "Encargo " & enc.Codigo & " não foi encontrado para o Produto selecionado.")
                            Exit Sub
                        End If
                    Next
                Else
                    MsgBox(Me.Page, "Encargos não encontrado para o Produto selecionado.")
                    Exit Sub
                End If

                CType(grdNFGProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text = objProduto.Codigo
                CType(grdNFGProdutos.Rows(index).FindControl("txtNomeProduto"), TextBox).Text = objProduto.Nome
                CType(grdNFGProdutos.Rows(index).FindControl("txtUnidade"), TextBox).Text = objProduto.Unidade

                CType(grdNFGProdutos.Rows(index).FindControl("txtFator"), TextBox).Text = 1
                If Not objProduto Is Nothing OrElse objProduto.UnidadesDeComercializacao.Count > 0 Then
                    CType(grdNFGProdutos.Rows(index).FindControl("txtFator"), TextBox).Text = objProduto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = objProduto.Unidade).First.FatorConversao
                End If

                CType(grdNFGProdutos.Rows(index).FindControl("txtNCM"), TextBox).Text = objProduto.NCM
                CType(grdNFGProdutos.Rows(index).FindControl("txtOperacaoEstado"), TextBox).Text = newObjOE.Codigo

                If Not objNotaFiscal.Itens(index).CodigoProduto = objProduto.Codigo Then
                    CType(grdNFGProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = True
                    CType(grdNFGProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Enabled = False
                End If

                CType(grdNFGProdutos.Rows(index).FindControl("chkNCM"), CheckBox).Enabled = False

                If Not objNotaFiscal.Itens(index).Produto.NCM = objProduto.NCM Then
                    CType(grdNFGProdutos.Rows(index).FindControl("chkNCM"), CheckBox).Checked = True
                End If

                Session.Remove("objProduto" & HID.Value)
            Else
                CType(grdNFGProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = False
                CType(grdNFGProdutos.Rows(index).FindControl("chkNCM"), CheckBox).Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNFConsultaNFG" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtCliente.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""

        grdNFGProdutos.DataSource = Nothing
        grdNFGProdutos.DataBind()

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)

        Try
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            SalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaNFG" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaNFG" & HID.Value), NotaFiscal))

                If objNotaFiscal.TotalDeNotasFiscais() Then
                    'CASO ENTRE NESSA MENSAGEM DEVAMOS ANALISAR O PEDIDO PARA PODER AJUSTAR - FURLAN - 15/01/2025
                    MsgBox(Me.Page, "Essa nota não pode ser corrigida por aqui pois existem mais notas nesse Pedido. Entre em Contato com o Suporte.", eTitulo.Info)
                    Exit Sub
                End If

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie

                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False

                grdNFGProdutos.DataSource = objNotaFiscal.Itens
                grdNFGProdutos.DataBind()

                SalvaNotaFiscal()
                Session.Remove("objNFConsultaNFG" & HID.Value)

                lnkConsultar.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampos() As Boolean
        For Each row As GridViewRow In grdNFGProdutos.Rows
            Dim chkProduto As CheckBox = CType(row.FindControl("chkProduto"), CheckBox)
            Dim chkObservacaoProd As CheckBox = CType(row.FindControl("chkObservacaoProd"), CheckBox)
            Dim chkNCM As CheckBox = CType(row.FindControl("chkNCM"), CheckBox)

            If chkProduto.Checked OrElse chkObservacaoProd.Checked OrElse chkNCM.Checked Then
                Return True
            End If
        Next

        MsgBox(Me.Page, "Não foi(ram) selecionado(s) item(s).")
        Return False
    End Function

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarProdutoNFG")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class

Public Class NotaParaAlteracao
    Public Property Nota As String
    Public Property EntradaSaida As String
    Public Property ProdutoAtual As String
    Public Property ProdutoCorrigido As String

    Public Sub New(nota As String, entradaSaida As String, produtoAtual As String, produtoCorrigido As String)
        Me.Nota = nota
        Me.EntradaSaida = entradaSaida
        Me.ProdutoAtual = produtoAtual
        Me.ProdutoCorrigido = produtoCorrigido
    End Sub

End Class