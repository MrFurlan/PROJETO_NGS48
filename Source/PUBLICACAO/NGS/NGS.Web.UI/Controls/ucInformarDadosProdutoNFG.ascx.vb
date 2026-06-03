Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml

Public Class ucInformarDadosProdutoNFG
    Inherits BaseUserControl

    Private _xmlDoc As XmlDocument

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

    Public Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
        Set(ByVal value As DataSet)
            If value Is Nothing Then
                Session.Remove("dsXml" & HID.Value)
            Else
                Session("dsXml" & HID.Value) = value
            End If
        End Set
    End Property

    Public Property SessaoXmlDoc() As XmlDocument
        Get
            Return CType(Session("XmlDoc" & HID.Value), XmlDocument)
        End Get
        Set(ByVal value As XmlDocument)
            If value Is Nothing Then
                Session.Remove("XmlDoc" & HID.Value)
            Else
                Session("XmlDoc" & HID.Value) = value
            End If
        End Set
    End Property

#End Region

#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "S", True)
            ddl.Carregar(cmbGrupoProdutoCusto, CarregarDDL.Tabela.GrupoProduto, "S", True)
            ddl.Carregar(cmbCentroCusto, CarregarDDL.Tabela.CentroDeCustoDescricao, "LEN(CentroDeCusto_Id) = 5")
        End If
        cmbGrupoProduto.Enabled = True
        cmbProdutos.Enabled = True
        cmbGrupoProdutoCusto.Enabled = True
        cmbProdutosCusto.Enabled = True
        BtnConsultaProduto.Visible = True
        'cmdConsultaOperacao.Visible = True
        imbOperacao.Visible = True
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
        ElseIf Session("objProdutoCustoNFG" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoCustoNFG" & HID.Value)
            Session.Remove("objProdutoCustoNFG" & HID.Value)
            cmbGrupoProdutoCusto.SelectedValue = objProduto.CodigoGrupo
            ddl.Carregar(cmbProdutosCusto, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProdutoCusto.SelectedValue & "'", True)
            cmbProdutosCusto.SelectedValue = objProduto.Codigo
            cmbProdutosCusto_SelectedIndexChanged(Nothing, Nothing)

            SessaoRecuperaProdutoNota()

            objProdutoNota.CodigoProdutoCusto = objProduto.Codigo

            SessaoSalvaProdutoNota()

        ElseIf Session("objOperacoesNFG" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            SessaoRecuperaProdutoNota()
            Dim objSubOperacao As [Lib].Negocio.SubOperacao = CType(obj, [Lib].Negocio.SubOperacao)
            Session.Remove("objOperacoesNFG" & HID.Value)

            objNotaFiscal.CodigoOperacao = objSubOperacao.CodigoOperacao
            objNotaFiscal.CodigoSubOperacao = objSubOperacao.Codigo
            objProdutoNota.CodigoOperacao = objSubOperacao.CodigoOperacao
            objProdutoNota.CodigoSubOperacao = objSubOperacao.Codigo

            txtOperacao.Text = String.Format("{0}-{1} - {2}", objSubOperacao.CodigoOperacao, objSubOperacao.Codigo, objSubOperacao.Descricao)
            SessaoSalvaProdutoNota()
            SessaoSalvaNotaFiscal()

        End If

    End Sub
#End Region

#Region "Botões"
    Protected Sub lnkGravar_Click(sender As Object, e As EventArgs) Handles lnkGravar.Click
        AdicionarProduto()
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divInformarDadosProdutoNFG")
    End Sub

    Protected Sub cmbGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGrupoProduto.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        If (objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.NotaDeProdutor AndAlso objNotaFiscal.Cliente.CodigoCategoria < 4) OrElse
            ((Left(objNotaFiscal.Empresa.Codigo, 8) = "05366261" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "38198213" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44005444" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "44979506" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "62747840" _
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "48984539") And (cmbGrupoProduto.SelectedValue = "10101" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10102" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "10103" _
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

        txtOperacao.Text = String.Empty

        SessaoRecuperaNotaFiscal()

        Dim i As Integer = 0

        SessaoRecuperaProdutoNota()

        If objProdutoNota Is Nothing Then
            objProdutoNota = New NotaFiscalXItem(objNotaFiscal)
        End If

        objProdutoNota.CodigoProduto = cmbProdutos.SelectedValue

        objProdutoNota.Sequencia = 1
        objProdutoNota.Produto = Nothing
        '#FimBaseDeCalculo
        'HIDBaseDeCalculo.Value = objProdutoNota.Produto.BaseCalculo
        objProdutoNota.Encargos = Nothing
        SessaoSalvaProdutoNota()
        'End If
    End Sub

    Protected Sub BtnConsultaProduto_Click(sender As Object, e As EventArgs) Handles BtnConsultaProduto.Click
        Dim ucConsultaProduto As ucConsultaProduto = Me.Page.FindControlRecursive("ucConsultaProduto")
        If ucConsultaProduto IsNot Nothing Then
            Session("Where" & HID.Value) = " Situacao = 1 And Agrupar = 'S' "
            ucConsultaProduto.Limpar()
            ucConsultaProduto.MainUserControl = Me
            'ucConsultaProduto.BuscarProduto()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoNFG" & HID.Value, txtNome.ClientID, True)
        End If
    End Sub

    Protected Sub BtnConsultaProdutoCusto_Click(sender As Object, e As EventArgs) Handles BtnConsultaProdutoCusto.Click
        Dim ucConsultaProduto As ucConsultaProduto = Me.Page.FindControlRecursive("ucConsultaProduto")
        If ucConsultaProduto IsNot Nothing Then
            Session("Where" & HID.Value) = " Situacao = 1 "
            ucConsultaProduto.Limpar()
            ucConsultaProduto.MainUserControl = Me
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoCustoNFG" & HID.Value, txtNome.ClientID, True)
        End If
    End Sub


    Protected Sub cmbGrupoProdutoCusto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGrupoProdutoCusto.SelectedIndexChanged
        ddl.Carregar(cmbProdutosCusto, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProdutoCusto.SelectedValue & "'", True)
    End Sub

    Protected Sub cmbProdutosCusto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProdutosCusto.SelectedIndexChanged

        SessaoRecuperaProdutoNota()

        objProdutoNota.CodigoProdutoCusto = cmbProdutosCusto.SelectedValue

        SessaoSalvaProdutoNota()

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

            If objNotaFiscal Is Nothing Then
                SessaoRecuperaNotaFiscal()
            End If

            If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso (objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                parameters("documento") = "CTE"
            Else
                parameters("documento") = ""
            End If

            parameters("codEmpresa") = Left(objNotaFiscal.CodigoEmpresa, 8)
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
            ucConsultaOperacoes.SetarHID(HID.Value)
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
        imbOperacao.Visible = True
        cmbProdutos.Items.Clear()
        cmbCentroCusto.SelectedIndex = 0
        divCentroDeCusto.Visible = False
        divProdutoCusto.Visible = False
        txtOperacao.Text = String.Empty
        PosicaoItem.Value = -1
        '#FimBaseDeCalculo
        'HIDBaseDeCalculo.Value = String.Empty
        If objNotaFiscal Is Nothing Then SessaoRecuperaNotaFiscal()
        objProdutoNota = New NotaFiscalXItem(objNotaFiscal)
        SessaoSalvaProdutoNota()

        HttpContext.Current.Session("ssMessage") = String.Empty

    End Sub

    Public Sub AdicionarProduto()

        SessaoRecuperaNotaFiscal()
        SessaoRecuperaProdutoNota()

        Try

            If ValidarCamposProdutos() Then

                If cmbCentroCusto.SelectedIndex > 0 Then
                    objProdutoNota.Encargos.EncProduto.CentroDeCusto = cmbCentroCusto.SelectedValue
                    objProdutoNota.CentroDeCustoInformado = cmbCentroCusto.SelectedValue
                    objProdutoNota.TemCentroDeCusto = True
                End If

                SessaoSalvaProdutoNota()
                SessaoSalvaNotaFiscal()
                CType(Me.Page, NotasFiscaisGerais).AtualizarDadosInformados()

                Popup.CloseDialog(Me.Page, "divInformarDadosProdutoNFG")

            End If

        Catch ex As Exception

            MsgBox(Me.Page, ex.Message)

        End Try

    End Sub

    Private Function ValidarCamposProdutos() As Boolean

        If cmbGrupoProduto.SelectedIndex < 1 Then
            MsgBox(Me.Page, "É necessário selecionar um grupo de produto!")
            Return False
        ElseIf cmbProdutos.SelectedIndex < 1 Then
            MsgBox(Me.Page, "É necessário selecionar um produto!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtOperacao.Text) Then
            MsgBox(Me.Page, "É necessário selecionar a operação!")
            Return False
        ElseIf objProdutoNota.CodigoOperacao = 0 Then
            MsgBox(Me.Page, "É necessário informar a operação!")
            Return False
        ElseIf objProdutoNota.CodigoSubOperacao = 0 Then
            MsgBox(Me.Page, "É necessário informar a suboperação!")
            Return False
        ElseIf objProdutoNota.Encargos.Count > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.ProdutoParaCusto AndAlso (objProdutoNota.CodigoProdutoCusto Is Nothing OrElse objProdutoNota.CodigoProdutoCusto.Length = 0) Then
            MsgBox(Me.Page, "É necessário informar produto para custo!")
            Return False
        End If

        Return True

    End Function
#End Region

End Class