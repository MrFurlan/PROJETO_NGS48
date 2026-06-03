Imports NGS.Lib.Negocio
Imports System.Threading
Imports System.IO

Public Class CCeProduto
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
                If Funcoes.VerificaPermissao("CCeProduto", "ACESSAR") Then
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
            If Funcoes.VerificaPermissao("CCeProduto", "LEITURA") Then
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

                Session("ssCampo" & HID.Value) = "NotaXItens"

                ucConsultaPedidosXNotas.BindGridView()
                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaCCe" & HID.Value.ToString)
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
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
            If Funcoes.VerificaPermissao("CCeProduto", "ALTERAR") Then
                RecuperaNotaFiscal()

                If ValidaCampos() Then

                    Dim Sqls As New ArrayList
                    Dim msgNFE As String = String.Empty
                    Dim Sql As String = String.Empty

                    For Each row In grdCceProdutos.Rows
                        If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                            Dim nI As Integer = CType(CType(row.FindControl("chkProduto"), CheckBox).NamingContainer, GridViewRow).RowIndex
                            If objNotaFiscal.Itens(nI).NotasDevolucao IsNot Nothing AndAlso objNotaFiscal.Itens(nI).NotasDevolucao.Count > 0 Then
                                MsgBox(Me.Page, "Produto vinculado a Devolução não pode ser alterado/corrigido.", eTitulo.Info)
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If

                            Sql = "SELECT NotaReferencial_Id " & vbCrLf & _
                                    "  FROM NotaFiscalReferencial" & vbCrLf & _
                                    " WHERE EmpresaReferencial_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                    "   AND EndEmpresaReferencial_Id   =" & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                    "   AND ClienteReferencial_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                    "   AND EndClienteReferencial_Id   =" & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                    "   AND EntradaSaidaReferencial_Id ='" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                                    "   AND SerieReferencial_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                                    "   AND NotaReferencial_Id         =" & objNotaFiscal.Codigo & vbCrLf & _
                                    "   AND ProdutoReferencial_Id      ='" & objNotaFiscal.Itens(nI).CodigoProduto & "'" & vbCrLf & _
                                    "   AND CFOPReferencial_Id         =" & objNotaFiscal.Itens(nI).CFOP & vbCrLf & _
                                    "   AND SequenciaReferencial_Id    =" & objNotaFiscal.Itens(nI).Sequencia

                            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotaFiscalReferencial1")

                            If ds IsNot Nothing AndAlso ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                                MsgBox(Me.Page, "Produto " & objNotaFiscal.Itens(nI).CodigoProduto & " com Nota Fiscal Referenciada não pode ser alterada/corrigida.")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If

                            Sql = "SELECT NotaReferencial_Id " & vbCrLf & _
                                    "  FROM NotaFiscalReferencial" & vbCrLf & _
                                    " WHERE Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                    "   AND EndEmpresa_Id   =" & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                    "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                    "   AND EndCliente_Id   =" & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                    "   AND EntradaSaida_Id ='" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                                    "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                                    "   AND Nota_Id         =" & objNotaFiscal.Codigo & vbCrLf & _
                                    "   AND Produto_Id      ='" & objNotaFiscal.Itens(nI).CodigoProduto & "'" & vbCrLf & _
                                    "   AND CFOP_Id         =" & objNotaFiscal.Itens(nI).CFOP & vbCrLf & _
                                    "   AND Sequencia_Id    =" & objNotaFiscal.Itens(nI).Sequencia

                            Dim ds2 As DataSet = Banco.ConsultaDataSet(Sql, "NotaFiscalReferencial2")

                            If ds2 IsNot Nothing AndAlso ds2.Tables(0) IsNot Nothing AndAlso ds2.Tables(0).Rows.Count > 0 Then
                                MsgBox(Me.Page, "Produto " & objNotaFiscal.Itens(nI).CodigoProduto & " com Nota Fiscal Referenciada não pode ser alterada/corrigida.")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If

                            If objNotaFiscal.Itens(nI).Lotes IsNot Nothing AndAlso objNotaFiscal.Itens(nI).Lotes.Count > 0 Then
                                For Each nXlt In objNotaFiscal.Itens(nI).Lotes
                                    Sql = "Select Ordem_Id FROM OrdemDeProducaoXConsumoXLote" & vbCrLf & _
                                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                            "  and Produto_Id      ='" & objNotaFiscal.Itens(nI).CodigoProduto & "'" & vbCrLf & _
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

                    If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                        If [Lib].Negocio.DocumentoEletronico.CCeProduto(objNotaFiscal, msgNFE, grdCceProdutos) Then
                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-impcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)
                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)
                        Else
                            MsgBox(Me.Page, "Erro na Carta de Correção: " & msgNFE & ". Tente novamente ou entre em contato com o Suporte.")
                            Exit Sub
                        End If
                    End If

                    Sql = " Update NotasXEmbalagens set " & vbCrLf & _
                             "      Quantidade            = " & Str(objNotaFiscal.Quantidade) & "," & vbCrLf & _
                             "      Especie               = '" & objNotaFiscal.Especie & "'," & vbCrLf & _
                             "      Marca                 = '" & objNotaFiscal.Marca & "'," & vbCrLf & _
                             "      Numero                = '" & objNotaFiscal.Numero & "'," & vbCrLf & _
                             "      PesoBruto             = " & Str(objNotaFiscal.PesoBruto) & "," & vbCrLf & _
                             "      PesoLiquido           = " & Str(objNotaFiscal.PesoLiquido) & vbCrLf & _
                             "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                             "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                             "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                             "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                             "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                             "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                             "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                    Sqls.Add(Sql)

                    Dim temProdutoPedido As Boolean = False
                    For Each item In objNotaFiscal.Pedido.Itens
                        If item.CodigoProduto = Trim(CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text) Then
                            temProdutoPedido = True
                            Exit For
                        End If
                    Next

                    For Each row In grdCceProdutos.Rows
                        If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                            index = CType(CType(row.FindControl("chkProduto"), CheckBox).NamingContainer, GridViewRow).RowIndex

                            'Incluir Produto no Pedido e Nota Fiscal
                            If Not temProdutoPedido Then
                                Sql = "insert into PedidoXItem (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Classificacao, Retencao, OperacaoXEstado, UnidadeComercializacao)" & vbCrLf & _
                                      "SELECT      Empresa_Id, EndEmpresa_Id, Pedido_Id, '" & CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', Classificacao, Retencao, " & CType(grdCceProdutos.Rows(index).FindControl("txtOperacaoEstado"), TextBox).Text & ", UnidadeComercializacao" & vbCrLf & _
                                      "FROM        PedidoXItem" & vbCrLf & _
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "insert into  PedidoXItemXLancamento (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, PedidoItem_Id, TipoDeLancamento, Movimento, Quantidade, UnitarioOficial, UnitarioMoeda, TotalOficial, " & vbCrLf & _
                                      "                      TotalMoeda, DataEntrega, UnitarioOficialCompra, UnitarioMoedaCompra, UsuarioLiberacao, UsuarioLiberacaoData, QuantidadeComercializacao)" & vbCrLf & _
                                      "SELECT     Empresa_Id, EndEmpresa_Id, Pedido_Id, '" & CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', PedidoItem_Id, TipoDeLancamento, Movimento, Quantidade, UnitarioOficial, UnitarioMoeda, TotalOficial, " & vbCrLf & _
                                      "                      TotalMoeda, DataEntrega, UnitarioOficialCompra, UnitarioMoedaCompra, UsuarioLiberacao, UsuarioLiberacaoData, QuantidadeComercializacao" & vbCrLf & _
                                      "FROM         PedidoXItemXLancamento" & vbCrLf & _
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "insert into PedidosXEncargos (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Encargo_Id, BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda)" & vbCrLf & _
                                      "SELECT      Empresa_Id, EndEmpresa_Id, Pedido_Id, '" & CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', Encargo_Id, BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda" & vbCrLf & _
                                      "FROM        PedidosXEncargos" & vbCrLf & _
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)
                            End If

                            Sql = "insert into NotasFiscaisXItens (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Pedido, Deposito, EndDeposito, DepositoTerceiro, " & vbCrLf & _
                            "                      EndDepositoTerceiro, PesoFiscal, QuantidadeFisica, QuantidadeFiscal, Unitario, Valor, PesoQuantidade, Operacao, SubOperacao, Lote, ObservacoesDoProduto, " & vbCrLf & _
                            "                      Classificacao, Sequencia_Id, Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeDeEmbalagem, Rateado, NumeroDeclaracaoImportacao, DataRegistroDI, " & vbCrLf & _
                            "                      LocalDesembarqueImportacao, EstadoDesembarqueImportacao, DataDesembarqueImportacao, Fabricante, EndFabricante, NumeroPecas, Fixacao, ValorMoeda, " & vbCrLf & _
                            "                      ValorLiquido, ValorLiquidoMoeda, OperacaoXEstado)" & vbCrLf & _
                            "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, '" & CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', CFOP_Id, Pedido, Deposito, EndDeposito, DepositoTerceiro, " & vbCrLf & _
                            "                      EndDepositoTerceiro, PesoFiscal, QuantidadeFisica, QuantidadeFiscal, Unitario, Valor, PesoQuantidade, Operacao, SubOperacao, Lote," & vbCrLf

                            If CType(row.FindControl("chkObservacaoProd"), CheckBox).Checked Then
                                Sql &= "'" & Funcoes.EliminarCaracteresEspeciaisNF(CType(grdCceProdutos.Rows(index).FindControl("txtObservacaoProd"), TextBox).Text) & "'," & vbCrLf
                            Else
                                Sql &= " ObservacoesDoProduto, " & vbCrLf
                            End If

                            Sql &= "                      Classificacao, Sequencia_Id, Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeDeEmbalagem, Rateado, NumeroDeclaracaoImportacao, DataRegistroDI, " & vbCrLf & _
                            "                      LocalDesembarqueImportacao, EstadoDesembarqueImportacao, DataDesembarqueImportacao, Fabricante, EndFabricante, NumeroPecas, Fixacao, ValorMoeda, " & vbCrLf & _
                            "                      ValorLiquido, ValorLiquidoMoeda, " & CType(grdCceProdutos.Rows(index).FindControl("txtOperacaoEstado"), TextBox).Text & vbCrLf & _
                            "FROM         NotasFiscaisXItens" & vbCrLf & _
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            Sql = "insert into NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Encargo_Id, Base, Percentual, Valor, " & vbCrLf & _
                            "                      CentroDeCusto, Sequencia_id, PercentualExibicao)" & vbCrLf & _
                            "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, '" & CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', CFOP_Id, Encargo_Id, Base, Percentual, Valor, " & vbCrLf & _
                            "                      CentroDeCusto, Sequencia_id, PercentualExibicao" & vbCrLf & _
                            "FROM         NotasFiscaisXEncargos" & vbCrLf & _
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            If objNotaFiscal.Itens(index).Lotes IsNot Nothing AndAlso objNotaFiscal.Itens(index).Lotes.Count > 0 Then
                                For Each nXl In objNotaFiscal.Itens(index).Lotes

                                    Sql = "insert into NotaFiscalXLote (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Lote_Id, Quantidade, Validade, QuantidadeDeConsumo)" & vbCrLf & _
                                    "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, '" & CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text & "', " & vbCrLf & _
                                    "       CFOP_Id, Sequencia_Id, Lote_Id, Quantidade, Validade, QuantidadeDeConsumo" & vbCrLf & _
                                    "FROM         NotaFiscalXLote" & vbCrLf & _
                                    "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                    "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                    "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                    "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                    "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                    "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                                    "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                                    "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                                    "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                                    "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'" & vbCrLf & _
                                    "  and Lote_Id         ='" & nXl.Lote & "'"
                                    Sqls.Add(Sql)

                                Next
                            End If


                            'Remover Produto da Nota Fiscal e do Pedido
                            Sql = "DELETE FROM NotasFiscaisXEncargos" & vbCrLf & _
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            If objNotaFiscal.Itens(index).Lotes IsNot Nothing AndAlso objNotaFiscal.Itens(index).Lotes.Count > 0 Then
                                For Each nXl In objNotaFiscal.Itens(index).Lotes

                                    Sql = "DELETE FROM NotaFiscalXLote" & vbCrLf & _
                                    "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                    "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                    "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                    "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                    "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                    "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                                    "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                                    "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                                    "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                                    "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'" & vbCrLf & _
                                    "  and Lote_Id         ='" & nXl.Lote & "'"
                                    Sqls.Add(Sql)

                                Next
                            End If

                            Sql = "DELETE FROM NotasFiscaisXItens" & vbCrLf & _
                            "Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                            "  and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                            "  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                            "  and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                            "  and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                            "  and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                            "  and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                            "  and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                            "  and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                            "  and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)

                            If Not temProdutoPedido Then
                                Sql = "DELETE FROM PedidosXEncargos" & vbCrLf & _
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "DELETE FROM PedidoXItemXLancamento" & vbCrLf & _
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)

                                Sql = "DELETE FROM        PedidoXItem" & vbCrLf & _
                                      "Where Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                      "  and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                      "  and Pedido_id     = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                                      "  and Produto_id    = '" & objNotaFiscal.Itens(index).CodigoProduto & "'"
                                Sqls.Add(Sql)
                            End If
                        ElseIf CType(row.FindControl("chkObservacaoProd"), CheckBox).Checked Then
                            index = CType(CType(row.FindControl("chkObservacaoProd"), CheckBox).NamingContainer, GridViewRow).RowIndex
                            Sql = " Update NotasFiscaisXItens set " & vbCrLf & _
                                 "      ObservacoesDoProduto = '" & Funcoes.EliminarCaracteresEspeciaisNF(CType(grdCceProdutos.Rows(index).FindControl("txtObservacaoProd"), TextBox).Text) & "'" & vbCrLf & _
                                 "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                 "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                 "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                 "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                 "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                 "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                                 "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                                 "    and Produto_Id      ='" & objNotaFiscal.Itens(index).CodigoProduto & "'" & vbCrLf & _
                                 "    and CFOP_Id         ='" & objNotaFiscal.Itens(index).CFOP & "'" & vbCrLf & _
                                 "    and Sequencia_Id    ='" & objNotaFiscal.Itens(index).Sequencia & "'"
                            Sqls.Add(Sql)
                        End If
                    Next

                    Dim obs As String = objNotaFiscal.ObservacoesControleInterno

                    If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Carta de correção Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                        Else
                            obs = "Carta de correção Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                        End If
                    Else
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Alteração do Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        Else
                            obs = "Alteração do Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        End If
                    End If

                    Sql = " Update NotasFiscais set " & vbCrLf & _
                         "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                         "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                         "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                         "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                         "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                         "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                         "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                         "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                         "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
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

                    For Each row In grdCceProdutos.Rows
                        If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                            objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)

                            objNotaFiscal.NotaFiscalOriginal.Razao.ExcluiContabilizacaoNotaSQL(Sqls)

                            objNotaFiscal.NotaFiscalOriginal.Razao.ContabilizarNotaSql(Sqls)

                            Exit For
                        End If
                    Next

                    If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                        Dim strChave As String = objNotaFiscal.ChaveNFE & "cce"
                        Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)
                        If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                            Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))

                            Dim fs As Stream = New FileStream(Server.MapPath("Files/") & Server.HtmlDecode(String.Format("{0}.pdf", strChave)), FileMode.Open, FileAccess.Read)
                            Dim br As New BinaryReader(fs)
                            Dim Arqbytes As Byte() = br.ReadBytes(fs.Length)

                            objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With { _
                                                     .IUD = "I", _
                                                     .Codigo = String.Empty, _
                                                     .Descricao = String.Format("{0}.pdf", strChave),
                                                     .Arquivo = Arqbytes})

                            For Each arq As [Lib].Negocio.Arquivo In objNotaFiscal.Arquivos
                                arq.IUD = "I"
                                arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                                arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                                arq.CodigoCliente = objNotaFiscal.CodigoCliente
                                arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                                arq.CodigoNota = objNotaFiscal.Codigo
                                arq.Serie = objNotaFiscal.Serie
                                arq.CodigoPedido = objNotaFiscal.CodigoPedido
                                arq.SalvarSql(Sqls)
                            Next

                            If Sqls.Count > 0 Then
                                If Banco.GravaBanco(Sqls) Then
                                    MsgBox(Me.Page, "Alterado com sucesso e Carta de correção inserida. Verifique o Extrato do Pedido se a alteração ficou de acordo e o Espelho da Nota Fiscal se os Valores e Contabilização estão corretos.", eTitulo.Sucess)

                                    Extrato.Emitir(Me.Page, False, objNotaFiscal.Pedido.CodigoEmpresa, objNotaFiscal.Pedido.EnderecoEmpresa, "T", _
                                         objNotaFiscal.Pedido.Codigo)

                                    Dim espelho As New NotaFiscalEspelho()
                                    espelho.ExibirEspelho(Me.Page, objNotaFiscal.NotaFiscalOriginal)

                                    limpar()
                                Else
                                    Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                                    MsgBox(Me.Page, Observacao)
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox(Me.Page, "Alterado com sucesso mais não foi possível inserir a Carta de correção. Verifique o Extrato do Pedido se a alteração ficou de acordo e o Espelho da Nota Fiscal se os Valores e Contabilização estão corretos.", eTitulo.Sucess)

                            Extrato.Emitir(Me.Page, False, objNotaFiscal.Pedido.CodigoEmpresa, objNotaFiscal.Pedido.EnderecoEmpresa, "T", _
                                          objNotaFiscal.Pedido.Codigo)

                            Dim espelho As New NotaFiscalEspelho()
                            espelho.ExibirEspelho(Me.Page, objNotaFiscal.NotaFiscalOriginal)

                            limpar()
                        End If
                    Else
                        If Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, "Produto alterado com sucesso. Verifique o Extrato do Pedido se a alteração ficou de acordo e o Espelho da Nota Fiscal se os Valores e Contabilização estão corretos.", eTitulo.Sucess)
                            Extrato.Emitir(Me.Page, False, objNotaFiscal.Pedido.CodigoEmpresa, objNotaFiscal.Pedido.EnderecoEmpresa, "T", _
                                                                      objNotaFiscal.Pedido.Codigo)

                            Dim espelho As New NotaFiscalEspelho()
                            espelho.ExibirEspelho(Me.Page, objNotaFiscal.NotaFiscalOriginal)

                            limpar()
                        Else
                            Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                            MsgBox(Me.Page, Observacao)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkReimpressao_Click(sender As Object, e As EventArgs) Handles lnkReimpressao.Click
        Try
            RecuperaNotaFiscal()

            Dim Sqls As New ArrayList
            Dim Observacao As String = String.Empty
            Dim strChave As String = objNotaFiscal.ChaveNFE & "cce"
            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)

            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))

                Sqls.Clear()

                Dim fs As Stream = New FileStream(Server.MapPath("Files/") & Server.HtmlDecode(String.Format("{0}.pdf", strChave)), FileMode.Open, FileAccess.Read)
                Dim br As New BinaryReader(fs)
                Dim Arqbytes As Byte() = br.ReadBytes(fs.Length)

                objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With { _
                                         .IUD = "I", _
                                         .Codigo = String.Empty, _
                                         .Descricao = String.Format("{0}.pdf", strChave),
                                         .Arquivo = Arqbytes})

                For Each arq As [Lib].Negocio.Arquivo In objNotaFiscal.Arquivos
                    arq.IUD = "I"
                    arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                    arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                    arq.CodigoCliente = objNotaFiscal.CodigoCliente
                    arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                    arq.CodigoNota = objNotaFiscal.Codigo
                    arq.Serie = objNotaFiscal.Serie
                    arq.CodigoPedido = objNotaFiscal.CodigoPedido
                    arq.SalvarSql(Sqls)
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, "Alterado com sucesso e Carta de correção inserida.", eTitulo.Sucess)
                        limpar()
                    Else
                        Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                        MsgBox(Me.Page, Observacao)
                        Exit Sub
                    End If
                End If
            Else
                MsgBox(Me.Page, "Carta de correção não foi encontrada.", eTitulo.Sucess)
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
                            If newObjOE.CodigoSTICMS = enc.SituacaoTributaria AndAlso _
                                newObjOE.CodigoSTIPI = enc.SituacaoTributariaIPI AndAlso _
                                newObjOE.CodigoSTPISCOFINS = enc.SituacaoTributariaPISCOFINS Then
                            ElseIf objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                                CType(grdCceProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = False
                                MsgBox(Me.Page, "Encargo " & enc.Codigo & " com Situação (Tributária, IPI ou PIS/COFINS) diferente do Produto selecionado.")
                                Exit Sub
                            End If
                        Else
                            CType(grdCceProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = False
                            MsgBox(Me.Page, "Encargo " & enc.Codigo & " não foi encontrado para o Produto selecionado.")
                            Exit Sub
                        End If
                    Next
                Else
                    MsgBox(Me.Page, "Encargos não encontrado para o Produto selecionado.")
                    Exit Sub
                End If

                If Not objNotaFiscal.Itens(index).Produto.Unidade = objProduto.Unidade Then

                    If (objNotaFiscal.Itens(index).Produto.Unidade = "KG" AndAlso objProduto.Unidade = "KGS") Or
                        (objNotaFiscal.Itens(index).Produto.Unidade = "KGS" AndAlso objProduto.Unidade = "KG") Then
                        'NÃO FAZ NADA, CONTINUE A ALTERAÇÃO
                    Else
                        MsgBox(Me.Page, "Unidade de Comercialização do Produto selecionado é diferente do Produto atual, entre em contato com o Suporte.")
                        Exit Sub
                    End If
                End If

                CType(grdCceProdutos.Rows(index).FindControl("txtProduto"), TextBox).Text = objProduto.Codigo
                CType(grdCceProdutos.Rows(index).FindControl("txtNomeProduto"), TextBox).Text = objProduto.Nome
                CType(grdCceProdutos.Rows(index).FindControl("txtUnidade"), TextBox).Text = objProduto.Unidade
                CType(grdCceProdutos.Rows(index).FindControl("txtNCM"), TextBox).Text = objProduto.NCM
                CType(grdCceProdutos.Rows(index).FindControl("txtOperacaoEstado"), TextBox).Text = newObjOE.Codigo

                If Not objNotaFiscal.Itens(index).CodigoProduto = objProduto.Codigo Then
                    CType(grdCceProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = True
                    CType(grdCceProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Enabled = False
                End If

                CType(grdCceProdutos.Rows(index).FindControl("chkNCM"), CheckBox).Enabled = False

                If Not objNotaFiscal.Itens(index).Produto.NCM = objProduto.NCM Then
                    CType(grdCceProdutos.Rows(index).FindControl("chkNCM"), CheckBox).Checked = True
                End If

                Session.Remove("objProduto" & HID.Value)
            Else
                CType(grdCceProdutos.Rows(index).FindControl("chkProduto"), CheckBox).Checked = False
                CType(grdCceProdutos.Rows(index).FindControl("chkNCM"), CheckBox).Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtCliente.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""

        grdCceProdutos.DataSource = Nothing
        grdCceProdutos.DataBind()

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkReimpressao.Parent.Visible = False

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
            If Not Session("objNFConsultaCCe" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaCCe" & HID.Value), NotaFiscal))

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

                If System.IO.File.Exists("C:\Alfasig\LeituraNFe\-danfes\" & objNotaFiscal.ChaveNFE & "cce.pdf") Then
                    lnkReimpressao.Parent.Visible = True
                Else
                    lnkReimpressao.Parent.Visible = False
                End If

                grdCceProdutos.DataSource = objNotaFiscal.Itens
                grdCceProdutos.DataBind()

                SalvaNotaFiscal()
                Session.Remove("objNFConsultaCCe" & HID.Value)

                lnkConsultar.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampos() As Boolean
        For Each row As GridViewRow In grdCceProdutos.Rows
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
            Funcoes.Ajuda(Me.Page, "CCeProduto")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class