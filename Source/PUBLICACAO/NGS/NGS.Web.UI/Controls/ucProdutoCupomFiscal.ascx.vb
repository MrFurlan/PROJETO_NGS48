Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml

Public Class ucProdutoCupomFiscal
    Inherits BaseUserControl

#Region "Sessão"
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvaProdutoNota()
        Session("objProdutoNota" & HID.Value) = objProdutoNota
    End Sub

    Private Sub SessaoRecuperaProdutoNota()
        objProdutoNota = CType(Session("objProdutoNota" & HID.Value), [Lib].Negocio.NotaFiscalXItem)
    End Sub
#End Region

#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "S", True)
        End If
        cmbGrupoProduto.Enabled = True
        cmbProdutos.Enabled = True
        BtnConsultaProduto.Visible = True
    End Sub

    Public Sub InicializarUC(pPosicaoItem As Integer)
        SessaoRecuperaNotaFiscal()
        Limpar()
        PosicaoItem.Value = pPosicaoItem
        Dim itemRemovido As Boolean = False

        If pPosicaoItem = -1 Then

            cmbGrupoProduto.Items.Clear()
            ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)

            SessaoSalvaProdutoNota()

            Dim ucConsultaProduto As ucConsultaProdutoCupomFiscal = Me.Page.FindControlRecursive("ucConsultaProdutoCupomFiscal")

            ucConsultaProduto.Limpar()
            ucConsultaProduto.MainUserControl = Me
            ucConsultaProduto.SetarHID(HID.Value)
            ucConsultaProduto.BuscarProduto()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProdutoCupomFiscal(Me.Page, "objProdutoNFG" & HID.Value, txtNome.ClientID, True)

            Exit Sub
        End If


        ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
        cmbProdutos.SelectedValue = objProdutoNota.CodigoProduto
    End Sub

    Public Sub BloqueioEdicaoProduto()
        lnkLimpar.Parent.Visible = False
        cmbGrupoProduto.Enabled = False
        cmbProdutos.Enabled = False
        BtnConsultaProduto.Visible = False
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub
#End Region

#Region "Variáveis Locais"
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objProdutoNota As [Lib].Negocio.NotaFiscalXItem
#End Region

#Region "Retorno UC"
    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objProdutoNFG" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoNFG" & HID.Value)
            Session.Remove("objProdutoNFG" & HID.Value)
            cmbGrupoProduto.SelectedValue = objProduto.CodigoGrupo
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
            cmbProdutos.SelectedValue = objProduto.Codigo
            cmbProdutos_SelectedIndexChanged(Nothing, Nothing)
            'If objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = objProduto.Codigo AndAlso s.IUD = "D").Count > 0 Then
            '    Dim index As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = objProduto.Codigo)
            '    InicializarUC(index)
            '    BloqueioEdicaoProduto()
            'End If
        ElseIf Session("objOperacoesNFG" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim objSubOperacao As [Lib].Negocio.SubOperacao = CType(obj, [Lib].Negocio.SubOperacao)
            Session.Remove("objOperacoesNFG" & HID.Value)

            Dim fCliente As String = objNotaFiscal.CodigoCliente
            Dim fEndCliente As Integer = objNotaFiscal.EnderecoCliente
            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then

                objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                    objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                    objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                    objNotaFiscal.CodigoCliente = objNotaFiscal.NotasTrocaOrigem(0).CodigoCliente
                    objNotaFiscal.EnderecoCliente = objNotaFiscal.NotasTrocaOrigem(0).EnderecoCliente
                End If
            End If

            If PosicaoItem.Value = -1 AndAlso objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count = 0 Then
                objNotaFiscal.CodigoOperacao = objSubOperacao.CodigoOperacao
                objNotaFiscal.CodigoSubOperacao = objSubOperacao.Codigo
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count > 0 Then
                If PosicaoItem.Value = -1 Or PosicaoItem.Value > 0 Then
                    If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoOperacao <> objSubOperacao.CodigoOperacao Or objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoSubOperacao <> objSubOperacao.Codigo Then
                        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).SubOperacao.Financeiro <> objSubOperacao.Financeiro Then
                            MsgBox(Me.Page, "Produtos com operações diferentes devem estar configurados iguais no financeiro!")
                            Exit Sub
                        End If

                        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Any(Function(s) s.SubOperacao.CodigoGrupoContas <> objSubOperacao.CodigoGrupoContas) Then
                            MsgBox(Me.Page, "Produtos com operações diferentes devem ter o mesmo grupo de contas!")
                            Exit Sub
                        End If
                    End If
                ElseIf objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count > 1 And PosicaoItem.Value = 0 Then
                    If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(1).CodigoOperacao <> objSubOperacao.CodigoOperacao Or objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(1).CodigoSubOperacao <> objSubOperacao.Codigo Then
                        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(1).SubOperacao.Financeiro <> objSubOperacao.Financeiro Then
                            MsgBox(Me.Page, "Produtos com operações diferentes devem estar configurados iguais no financeiro!")
                            Exit Sub
                        End If

                        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Any(Function(s) s.SubOperacao.CodigoGrupoContas <> objSubOperacao.CodigoGrupoContas) Then
                            MsgBox(Me.Page, "Produtos com operações diferentes devem ter o mesmo grupo de contas!")
                            Exit Sub
                        End If
                    End If
                End If
            End If

            SessaoRecuperaProdutoNota()

            objProdutoNota.CodigoOperacao = objSubOperacao.CodigoOperacao
            objProdutoNota.CodigoSubOperacao = objSubOperacao.Codigo

            'Serve para Lista dos Encargos, gerar uma nova ao invés de tentar buscar o que está gravado na nota - Furlan - 25-01-2014
            'Na gravação CodigoTipoDeDocumentoFrete voltar a ser 0
            If Not objProdutoNota.Encargos Is Nothing AndAlso objProdutoNota.Encargos.Count = 0 Then objProdutoNota.Encargos = Nothing

            If objProdutoNota.Encargos Is Nothing OrElse objProdutoNota.Encargos.Count = 0 Then

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                    objNotaFiscal.TipoDeDocumentoFrete = Nothing

                    objNotaFiscal.CodigoCliente = fCliente
                    objNotaFiscal.EnderecoCliente = fEndCliente
                End If

                SessaoSalvaProdutoNota()
                SessaoSalvaNotaFiscal()

                MsgBox(Me.Page, "Não foram encontrados encargos do produto " & objProdutoNota.CodigoProduto & "-" & objProdutoNota.Produto.Descricao & " para a operação " & objSubOperacao.CodigoOperacao & "-" & objSubOperacao.Codigo)

                Exit Sub
            End If

            If Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto Then
                objProdutoNota.CarregandoEncargos = True
                objProdutoNota.TemCentroDeCusto = objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
                Dim cc As String = objProdutoNota.Encargos.EncProduto.CentroDeCusto
                If Not String.IsNullOrWhiteSpace(cc) AndAlso cc <> "0" Then
                    objProdutoNota.Encargos = Nothing
                    objProdutoNota.Encargos.EncProduto.CentroDeCusto = cc
                    objProdutoNota.CarregandoEncargos = False
                    objProdutoNota.Encargos.AtualizaLiquido()
                End If
            Else
                objProdutoNota.Encargos = Nothing
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                objNotaFiscal.TipoDeDocumentoFrete = Nothing

                objNotaFiscal.CodigoCliente = fCliente
                objNotaFiscal.EnderecoCliente = fEndCliente
            End If

            SessaoSalvaProdutoNota()
            SessaoSalvaNotaFiscal()
        End If
    End Sub
#End Region

#Region "Botões"
    Protected Sub lnkGravar_Click(sender As Object, e As EventArgs) Handles lnkGravar.Click
        SessaoRecuperaNotaFiscal()
        SessaoRecuperaProdutoNota()

        Try
            If ValidarCamposProdutos() Then

                txtValorTotal.Text = Math.Round((CDbl(txtQuantidade.Text) * CDbl(txtUnitario.Text)), 2).ToString("N2")

                objNotaFiscal.CodigoOperacao = 21
                objNotaFiscal.CodigoSubOperacao = 35

                objProdutoNota.CodigoOperacao = 21
                objProdutoNota.CodigoSubOperacao = 35

                objProdutoNota.PesoQuantidade = objProdutoNota.Produto.PesoQuantidade

                objProdutoNota.NotaFiscal.Cliente = New Cliente()
                objProdutoNota.NotaFiscal.Cliente.CodigoEstado = "MS"
                objProdutoNota.NotaFiscal.Cliente.InscricaoEstadual = "ISENTO"

                objProdutoNota.QuantidadeFiscal = txtQuantidade.Text
                objProdutoNota.QuantidadeFisica = txtQuantidade.Text
                objProdutoNota.Unitario = txtUnitario.Text

                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objProdutoNota.CodigoProduto)
                SaldoProdutoEstoque.CarregarResumoSaldoEmEstoque(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.SubOperacao)

                If SaldoProdutoEstoque.SaldoFiscal = 0 Then
                    MsgBox(Me.Page, "Produto sem estoque disponível", eTitulo.Info)
                Else
                    SessaoSalvaProdutoNota()
                    SessaoSalvaNotaFiscal()

                    Limpar()

                    CType(Me.Page, CupomFiscal).AtualizarItensNoGrid()

                End If

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)

            Limpar()

            Popup.CloseDialog(Me.Page, "divProdutoNFG")

            CType(Me.Page, NotasFiscaisGerais).LimparCampos()
        End Try

    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divProdutoCupomFiscal")
    End Sub

    Protected Sub cmbGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGrupoProduto.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        If (objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.NotaDeProdutor AndAlso objNotaFiscal.Cliente.CodigoCategoria < 4) OrElse
            ((Left(objNotaFiscal.Empresa.Codigo, 8) = "05366261" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "38198213" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62780383" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "63358210" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44005444" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44979506" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "48984539") And (cmbGrupoProduto.SelectedValue = "10101" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10102" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10401" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "30101" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10110" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10111" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10105" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10106" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10112" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "19101" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "40201")) Then
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
        Else
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "' AND Agrupar = 'S' ", True)
        End If
    End Sub

    Protected Sub cmbProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProdutos.SelectedIndexChanged
        txtQuantidade.Text = String.Empty

        txtUnitario.Text = String.Empty
        txtValorTotal.Text = String.Empty

        SessaoRecuperaNotaFiscal()

        Dim i As Integer = 0

        If Not objNotaFiscal.Itens Is Nothing Then i = objNotaFiscal.Itens.FindIndex(Function(s) s.IUD <> "D" And s.CodigoProduto = cmbProdutos.SelectedValue)

        'If objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = cmbProdutos.SelectedValue AndAlso s.IUD = "D").Count > 0 Then
        '    Dim index As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = cmbProdutos.SelectedValue)
        '    InicializarUC(index)
        '    BloqueioEdicaoProduto()
        'Else -> une com o if abaixo elseif
        If i >= 0 Then
            InicializarUC(i)
        Else
            SessaoRecuperaProdutoNota()
            objProdutoNota.CodigoProduto = cmbProdutos.SelectedValue
            objProdutoNota.Produto = Nothing
            '#FimBaseDeCalculo
            'HIDBaseDeCalculo.Value = objProdutoNota.Produto.BaseCalculo
            objProdutoNota.Encargos = Nothing
            SessaoSalvaProdutoNota()
        End If
    End Sub

    Protected Sub BtnConsultaProduto_Click(sender As Object, e As EventArgs) Handles BtnConsultaProduto.Click
        Dim ucConsultaProduto As ucConsultaProduto = Me.Page.FindControlRecursive("ucConsultaProduto")
        If ucConsultaProduto IsNot Nothing Then
            Session("Where" & HID.Value) = " Situacao = 1 And Agrupar = 'S' "
            ucConsultaProduto.Limpar()
            ucConsultaProduto.MainUserControl = Me
            ucConsultaProduto.BuscarProduto()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoNFG" & HID.Value, txtNome.ClientID, True)
        End If
    End Sub

    Protected Sub imbOperacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        If String.IsNullOrWhiteSpace(cmbProdutos.SelectedValue) Then
            MsgBox(Me.Page, "Produto não foi selecionado.", eTitulo.Info)
            Exit Sub
        End If

        SessaoRecuperaNotaFiscal()
        Dim ucConsultaOperacoes As ucConsultaOperacoes = Me.Page.FindControlRecursive("ucConsultaOperacoes")
        If ucConsultaOperacoes IsNot Nothing Then
            HttpContext.Current.Session("ssCampo") = "Livre"
            Dim parameters As New Dictionary(Of String, Object)
            parameters("tipo") = "NFG"
            parameters("documento") = ""
            parameters("vigencia") = objNotaFiscal.Movimento.ToString("yyyy-MM-dd")
            parameters("prod") = cmbProdutos.SelectedValue
            If objNotaFiscal.Empresa IsNot Nothing Then
                parameters("ufOri") = objNotaFiscal.Empresa.CodigoEstado
            End If
            If objNotaFiscal.Cliente IsNot Nothing Then
                parameters("cliDes") = objNotaFiscal.Cliente.Codigo & "-" & objNotaFiscal.Cliente.CodigoEndereco
                If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                    objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                    objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                    parameters("cliDes") = objNotaFiscal.NotasTrocaOrigem(0).Cliente.Codigo & "-" & objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEndereco
                End If
            End If
            ucConsultaOperacoes.Limpar()
            ucConsultaOperacoes.MainUserControl = Me
            ucConsultaOperacoes.BindGridView(parameters)
            Popup.ConsultaDeOperacoes(Me.Page, "objOperacoesNFG" & HID.Value)
        End If
    End Sub

#End Region

#Region "Methods & Functions"
    Public Overrides Sub Limpar()
        If cmbGrupoProduto.Items.Count > 0 Then
            cmbGrupoProduto.SelectedIndex = 0
        End If
        lnkLimpar.Parent.Visible = True
        cmbGrupoProduto.Enabled = True
        cmbProdutos.Enabled = True
        BtnConsultaProduto.Visible = True
        cmbProdutos.Items.Clear()
        txtQuantidade.Text = String.Empty
        txtUnitario.Text = String.Empty
        txtValorTotal.Text = String.Empty
        txtValorTotalOld.Value = "0"
        PosicaoItem.Value = -1
        '#FimBaseDeCalculo
        'HIDBaseDeCalculo.Value = String.Empty
        If objNotaFiscal Is Nothing Then SessaoRecuperaNotaFiscal()
        objProdutoNota = New NotaFiscalXItem(objNotaFiscal)
        SessaoSalvaProdutoNota()
    End Sub

    Private Function ValidarCamposProdutos() As Boolean
        If cmbGrupoProduto.SelectedIndex < 1 Then
            MsgBox(Me.Page, "É necessário selecionar um grupo de produto!")
            Return False
        ElseIf cmbProdutos.SelectedIndex < 1 Then
            MsgBox(Me.Page, "É necessário selecionar um produto!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtQuantidade.Text) Then
            MsgBox(Me.Page, "É necessário informar a quantidade!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtUnitario.Text) Then
            MsgBox(Me.Page, "É necessário informar o valor unitário do produto!")
            Return False
        End If

        Return True
    End Function
#End Region

End Class