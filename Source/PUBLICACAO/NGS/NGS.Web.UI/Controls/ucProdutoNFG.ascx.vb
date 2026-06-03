Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml

Public Class ucProdutoNFG
    Inherits BaseUserControl

    Private _bReaproveitarDados As Boolean

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

    Private Sub SessaoSalvaProdutoNotaOrigem()
        Session("objProdutoNotaOrigem" & HID.Value) = objProdutoNotaOrigem
    End Sub

    Private Sub SessaoRecuperaProdutoNotaOrigem()
        objProdutoNotaOrigem = CType(Session("objProdutoNotaOrigem" & HID.Value), [Lib].Negocio.NotaFiscalXItem)
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

    Public Property ReaproveitarDados As Boolean
        Get
            Return hidReabroveitarDados.Value
        End Get
        Set(value As Boolean)
            hidReabroveitarDados.Value = value
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

    Public Sub InicializarUC(pPosicaoItem As Integer, ByVal bXMLAutomatico As Boolean, ByVal bReaproveitarDados As Boolean)

        SessaoRecuperaNotaFiscal()
        Limpar()
        PosicaoItem.Value = pPosicaoItem
        Dim itemRemovido As Boolean = False
        Dim produtoConsulta As New Produto()

        ReaproveitarDados = bReaproveitarDados

        If bReaproveitarDados AndAlso objNotaFiscal.Itens.Count > 0 AndAlso SessaoDsXML Is Nothing Then

            If pPosicaoItem = -1 Then
                objProdutoNota = objNotaFiscal.Itens.FirstOrDefault()
            Else
                objProdutoNota = objNotaFiscal.Itens(pPosicaoItem)
            End If

            txtOperacao.Text = objProdutoNota.CodigoOperacao & " - " & objProdutoNota.CodigoSubOperacao & "  " & objProdutoNota.SubOperacao.Descricao
            imbOperacao.Visible = True

            With objProdutoNota

                .IUD = "I"
                .CarregandoEncargos = True
                .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objProdutoNota, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                .CarregandoEncargos = False

                If Not .Encargos Is Nothing AndAlso .Encargos.Count() > 0 Then
                    .TemCentroDeCusto = .Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
                    objProdutoNota.CarregandoEncargos = False
                End If

            End With

            ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "S", True)

            cmbGrupoProduto.SelectedValue = objProdutoNota.Produto.CodigoGrupo
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & objProdutoNota.Produto.CodigoGrupo & "'", True)
            cmbProdutos.SelectedValue = objProdutoNota.Produto.Codigo

            If objProdutoNota.Encargos.Count > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing Then
                ddl.Carregar(cmbCentroCusto, CarregarDDL.Tabela.CentroDeCustoDescricao, "LEN(CentroDeCusto_Id) = 5")
                divCentroDeCusto.Visible = objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
                If Not objProdutoNota.CodigoProdutoCusto Is Nothing AndAlso objProdutoNota.CodigoProdutoCusto.Length > 0 Then
                    cmbCentroCusto.SelectedValue = objProdutoNota.CodigoProdutoCusto
                End If
            End If

            txtQuantidade.Enabled = True
            txtUnitario.Enabled = True

            If pPosicaoItem <> -1 Then
                txtQuantidade.Text = objProdutoNota.QuantidadeFiscal
                txtUnitario.Text = objProdutoNota.Unitario.ToString("N10")
                txtValorTotal.Text = objProdutoNota.ValorTotal.ToString("N2")
            End If

            SessaoSalvaProdutoNota()

            Exit Sub

        ElseIf pPosicaoItem = -1 Then

            If objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.NotaDeProdutor AndAlso objNotaFiscal.Cliente.CodigoCategoria < 4 Then
                cmbGrupoProduto.Items.Clear()

                ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)
            End If

            objProdutoNota.IUD = "I"
            SessaoSalvaProdutoNota()

            If objNotaFiscal.ChaveNFE.Length = 44 AndAlso objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.Nota Then
                Dim pNomeArquivo As String = String.Format("{0}-nfe", objNotaFiscal.ChaveNFE)

                'If System.IO.File.Exists(Server.MapPath(String.Format("~/Files/{0}.xml", pNomeArquivo))) Then
                '    Dim DsXml As New DataSet
                '    DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}.xml", pNomeArquivo)))

                '    If DsXml.Tables("prod").Rows(0)("NCM").ToString().Length > 0 Then

                '        Dim ucConsultaProduto As ucConsultaProduto = Me.Page.FindControlRecursive("ucConsultaProduto")
                '        If ucConsultaProduto IsNot Nothing Then
                '            Session("Where" & HID.Value) = " Agrupar = 'S' AND NCM = '" & DsXml.Tables("prod").Rows(0)("NCM").ToString() & "'"
                '            ucConsultaProduto.Limpar()
                '            ucConsultaProduto.MainUserControl = Me
                '            ucConsultaProduto.SetarHID(HID.Value)
                '            ucConsultaProduto.BuscarProduto(True)
                '            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
                '            Popup.ConsultaDeProduto(Me.Page, "objProdutoNFG" & HID.Value, txtNome.ClientID, True)
                '        End If
                '    End If
                'End If

                Exit Sub

            Else

                Dim iOperacaoFrete As Integer
                Dim iSubOperacaoFrete As Integer
                Dim sProdutoFrete As String = objNotaFiscal.Empresa.Empresa.CodigoProdutoDeFrete
                Dim produtoNosso As New Produto

                produtoNosso = New Produto(sProdutoFrete)

                'Se tiver'

                If produtoNosso.Codigo.Length > 0 Then

                    Dim historicoOP As ListClientesXHistoricoOperacoes = New ListClientesXHistoricoOperacoes("", 0, produtoNosso.Codigo, 0, 0, objNotaFiscal.CodigoTipoDeDocumento)

                    If historicoOP.Count > 0 Then

                        If historicoOP.Where(Function(x) x.Cliente_Id = objNotaFiscal.CodigoCliente And x.TipoDocumento = objNotaFiscal.CodigoTipoDeDocumento).Count() > 0 Then
                            iOperacaoFrete = historicoOP.Where(Function(x) x.Cliente_Id = objNotaFiscal.CodigoCliente And x.TipoDocumento = objNotaFiscal.CodigoTipoDeDocumento).FirstOrDefault().Operacao_Id
                            iSubOperacaoFrete = historicoOP.Where(Function(x) x.Cliente_Id = objNotaFiscal.CodigoCliente And x.TipoDocumento = objNotaFiscal.CodigoTipoDeDocumento).FirstOrDefault().SubOperacao_Id
                        Else
                            iOperacaoFrete = historicoOP(0).Operacao_Id
                            iSubOperacaoFrete = historicoOP(0).SubOperacao_Id
                        End If

                    End If

                End If

                If bXMLAutomatico Then

                    Dim sProdutoParaCusto As String = "10101"

                    If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" And iSubOperacaoFrete = 0 Then
                        iOperacaoFrete = 80
                        iSubOperacaoFrete = 2
                    End If

                    Dim fCliente As String = objNotaFiscal.CodigoCliente
                    Dim fEndCliente As Integer = objNotaFiscal.EnderecoCliente

encargoRT:
                    produtoConsulta = New Produto(sProdutoFrete)

                    ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "S", True)
                    ddl.Carregar(cmbGrupoProdutoCusto, CarregarDDL.Tabela.GrupoProduto, "S", True)
                    ddl.Carregar(cmbCentroCusto, CarregarDDL.Tabela.CentroDeCustoDescricao, "LEN(CentroDeCusto_Id) = 5")

                    If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then

                        objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                        If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                                    objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                            objNotaFiscal.CodigoCliente = objNotaFiscal.NotasTrocaOrigem(0).CodigoCliente
                            objNotaFiscal.EnderecoCliente = objNotaFiscal.NotasTrocaOrigem(0).EnderecoCliente
                        End If

                    End If


                    Dim Parametros As New OperacaoXEstado
                    Parametros.Empresa = Left(objNotaFiscal.Empresa.Codigo, 8)
                    Parametros.CodigoGrupoProduto = produtoConsulta.CodigoGrupo
                    Parametros.CodigoProduto = produtoConsulta.Codigo
                    Parametros.CodigoOperacao = iOperacaoFrete
                    Parametros.CodigoSubOperacao = iSubOperacaoFrete
                    Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado
                    Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado
                    Parametros.InicioVigencia = objNotaFiscal.Movimento
                    Dim OXE As New OperacaoXEstado(Parametros)

                    objNotaFiscal.CodigoOperacao = iOperacaoFrete

                    If iSubOperacaoFrete > 0 Then
                        objNotaFiscal.CodigoSubOperacao = iSubOperacaoFrete
                    End If

                    objNotaFiscal.NossaEmissao = False
                    objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                    cmbGrupoProduto.SelectedValue = produtoConsulta.CodigoGrupo
                    ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & produtoConsulta.CodigoGrupo & "'", True)
                    cmbProdutos.SelectedValue = produtoConsulta.Codigo

                    If objNotaFiscal.Itens.Count = 0 Then

                        objProdutoNota = New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)
                        With objProdutoNota

                            .IUD = "I"
                            .CodigoProduto = produtoConsulta.Codigo
                            .PesoQuantidade = produtoConsulta.PesoQuantidade
                            .CodigoOperacao = iOperacaoFrete
                            .CodigoSubOperacao = iSubOperacaoFrete
                            .CodigoOperacaoEstado = OXE.Codigo

                            If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" Then

                                objProdutoNota.CarregandoEncargos = True
                                .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objProdutoNota, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})

                                If Not .Encargos Is Nothing AndAlso .Encargos.Count() > 0 Then
                                    .TemCentroDeCusto = .Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
                                    objProdutoNota.CarregandoEncargos = False
                                End If

                            End If

                        End With

                        'objNotaFiscal.Itens.Add(objProdutoNota)

                    Else

                        objProdutoNota = objNotaFiscal.Itens.Where(Function(x) x.CodigoProduto = produtoConsulta.Codigo).FirstOrDefault()

                        If objProdutoNota Is Nothing Then
                            objProdutoNota = objNotaFiscal.Itens.FirstOrDefault()
                        End If

                    End If

                    If Not objNotaFiscal.SubOperacao Is Nothing Then
                        txtOperacao.Text = objNotaFiscal.CodigoOperacao & " - " & objNotaFiscal.CodigoSubOperacao & "  " & objNotaFiscal.SubOperacao.Descricao
                    End If

                    If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso iSubOperacaoFrete = 2 AndAlso (objProdutoNota.Encargos.Count() = 0 OrElse objProdutoNota.Encargos Is Nothing) Then
                        iSubOperacaoFrete = 1
                        GoTo encargoRT
                    End If

                    If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then

                        objNotaFiscal.CodigoCliente = fCliente
                        objNotaFiscal.EnderecoCliente = fEndCliente

                    End If

                    If objProdutoNota.Encargos Is Nothing OrElse objProdutoNota.Encargos.Count() = 0 And iSubOperacaoFrete > 0 Then
                        MsgBox(Me.Page, String.Format("Encargos não encontrados para a operação: {0}-{1} - Origem: {2} - Destino: {3}!", iOperacaoFrete, iSubOperacaoFrete, objNotaFiscal.Empresa.CodigoEstado, objNotaFiscal.Cliente.CodigoEstado))
                        'Exit Sub
                    End If

                    If objProdutoNota.Encargos.Count() > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing Then

                        divCentroDeCusto.Visible = objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
                        cmbCentroCusto.SelectedValue = sProdutoParaCusto

                    End If

                    Dim dValorTotal As Decimal

                    If Not SessaoDsXML Is Nothing Then

                        Dim dsXml As DataSet = SessaoDsXML

                        If Not dsXml.Tables("vPrest") Is Nothing Then

                            dValorTotal = dsXml.Tables("vPrest").Rows(0).Item("vRec").ToString.Replace(".", ",")

                        End If

                    End If

                    If objNotaFiscal.TipoDeDocumento.Codigo <> eTipoDeDocumento.CT_E_TOM Then

                        If objNotaFiscal.NotasTrocaOrigem Is Nothing OrElse objNotaFiscal.NotasTrocaOrigem.Count = 0 Then
                            Throw New Exception("Nota de origem não encontrada!")
                        End If

                        Dim notaOrigem As New [Lib].Negocio.NotaFiscal(objNotaFiscal.NotasTrocaOrigem.Item(0))
                        objProdutoNota.QuantidadeFiscal = notaOrigem.Itens(0).QuantidadeFiscal * IIf(notaOrigem.Itens(0).Produto.UnidadesDeComercializacao(0).CodigoUnidade <> "KG", 1000, notaOrigem.Itens(0).Produto.UnidadesDeComercializacao(0).FatorConversao)

                    End If

                    'ACRESCENTADO PARA IR NO ESTOQUE FÍSICO CASO ESTEJA MARCADO NA OPERAÇÃO - FURLAN - 30/07/2024
                    If Not objProdutoNota.SubOperacao Is Nothing AndAlso Not objProdutoNota.SubOperacao.EstoqueFisico Then
                        objProdutoNota.QuantidadeFisica = 0
                    End If

                    If objProdutoNota.QuantidadeFiscal = 0 And objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E_TOM Then
                        objProdutoNota.QuantidadeFiscal = 1
                    End If

                    objProdutoNota.ValorTotal = dValorTotal
                    objProdutoNota.Unitario = objProdutoNota.ValorTotal / objProdutoNota.QuantidadeFiscal

                    SessaoSalvaProdutoNota()

                Else

                    objNotaFiscal.NossaEmissao = False
                    objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                End If

            End If

        Else
            objProdutoNota = objNotaFiscal.Itens(pPosicaoItem)
            If Not objProdutoNota.IUD = "I" Then objProdutoNota.IUD = "U"

            '#FimBaseDeCalculo
            'HIDBaseDeCalculo.Value = objProdutoNota.Produto.BaseCalculo
            SessaoSalvaProdutoNota()
        End If

        If Not produtoConsulta Is Nothing AndAlso Not produtoConsulta.Codigo Is Nothing AndAlso produtoConsulta.Codigo.Length > 0 Then

            cmbGrupoProduto.SelectedValue = produtoConsulta.CodigoGrupo
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & produtoConsulta.CodigoGrupo & "'", True)
            cmbProdutos.SelectedValue = produtoConsulta.Codigo

        ElseIf Not objProdutoNota Is Nothing AndAlso Not objProdutoNota.Produto Is Nothing Then

            cmbGrupoProduto.SelectedValue = objProdutoNota.Produto.CodigoGrupo
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
            cmbProdutos.SelectedValue = objProdutoNota.CodigoProduto

            If Not itemRemovido Then txtOperacao.Text = objProdutoNota.CodigoOperacao & " - " &
            objProdutoNota.CodigoSubOperacao & "  " & objProdutoNota.SubOperacao.Descricao

            If objProdutoNota.Encargos.Count > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing Then
                divCentroDeCusto.Visible = objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
            End If

            If objProdutoNota.Encargos.Count > 0 AndAlso objProdutoNota.IUD = "U" Then
                If Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.ProdutoParaCusto AndAlso objProdutoNota.CodigoProdutoCusto.Length > 0 Then
                    divProdutoCusto.Visible = True

                    Dim objProduto = New [Lib].Negocio.Produto(objProdutoNota.CodigoProdutoCusto)

                    cmbGrupoProdutoCusto.SelectedValue = objProduto.CodigoGrupo
                    ddl.Carregar(cmbProdutosCusto, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProdutoCusto.SelectedValue & "'", True)
                    cmbProdutosCusto.SelectedValue = objProduto.Codigo
                End If
            End If

            If Not objProdutoNota.Encargos Is Nothing AndAlso
                   objProdutoNota.Encargos.Count > 0 AndAlso
                   Not objProdutoNota.Encargos.EncProduto Is Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(objProdutoNota.Encargos.EncProduto.CentroDeCusto) AndAlso
                   CInt(objProdutoNota.Encargos.EncProduto.CentroDeCusto) > 0 Then
                cmbCentroCusto.SelectedValue = objProdutoNota.Encargos.EncProduto.CentroDeCusto
            End If

            txtQuantidade.Text = objProdutoNota.QuantidadeFiscal.ToString("N4")
            If objProdutoNota.CodigoProduto <> objNotaFiscal.Empresa.Empresa.CodigoProdutoDeFrete And
                objProdutoNota.CodigoProduto <> objNotaFiscal.Empresa.Empresa.CodigoProdutoDeEstadia And
                objProdutoNota.CodigoProduto <> objNotaFiscal.Empresa.Empresa.CodigoProdutoDeMDFe Then
                txtQuantidade.Enabled = True
            End If

        End If

        If SessaoDsXML Is Nothing Then
            txtQuantidade.Enabled = True
            txtUnitario.Enabled = True
        End If

        txtQuantidade.Text = objProdutoNota.QuantidadeFiscal
        txtUnitario.Text = objProdutoNota.Unitario.ToString("N10")
        txtValorTotal.Text = objProdutoNota.ValorTotal.ToString("N2")
        txtValorTotalOld.Value = objProdutoNota.ValorTotal
        SetarMsgRetencao()

        If objProdutoNota.QuantidadeFiscal <> 0 Then
            BloqueioEdicaoProdutoXML()
        End If

        If objNotaFiscal.IUD = "U" And objNotaFiscal.Arquivos.Count() = 0 And SessaoDsXML Is Nothing Then
            txtQuantidade.Enabled = True
            txtUnitario.Enabled = True
        End If

        If pPosicaoItem <> -1 Then
            objProdutoNotaOrigem = objNotaFiscal.Itens(pPosicaoItem).Clone()
            If objProdutoNotaOrigem.CodigoProdutoOld Is Nothing OrElse objProdutoNotaOrigem.CodigoProdutoOld.Length = 0 Then
                objProdutoNotaOrigem.CodigoProdutoOld = objProdutoNotaOrigem.CodigoProduto
                objNotaFiscal.Itens(pPosicaoItem).CodigoProdutoOld = objProdutoNotaOrigem.CodigoProduto
            End If
            SessaoSalvaProdutoNotaOrigem()
        End If

    End Sub

    Public Sub InicializarUC(pPosicaoItem As Integer, dsXml As DataSet)

        SessaoRecuperaNotaFiscal()
        Limpar()
        PosicaoItem.Value = pPosicaoItem
        Dim itemRemovido As Boolean = False

        If pPosicaoItem = -1 Then

            If objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.NotaDeProdutor AndAlso objNotaFiscal.Cliente.CodigoCategoria < 4 Then
                cmbGrupoProduto.Items.Clear()

                ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)
            End If

            objProdutoNota.IUD = "I"

            txtValorTotalOld.Value = objProdutoNota.ValorTotal

            SessaoSalvaProdutoNota()

            'If objNotaFiscal.ChaveNFE.Length = 44 AndAlso objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.Nota Then
            '    Dim pNomeArquivo As String = String.Format("{0}-nfe", objNotaFiscal.ChaveNFE)

            '    If System.IO.File.Exists(Server.MapPath(String.Format("~/Files/{0}.xml", pNomeArquivo))) Then
            '        Dim DsXml As New DataSet
            '        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}.xml", pNomeArquivo)))

            '        If DsXml.Tables("prod").Rows(0)("NCM").ToString().Length > 0 Then

            '            Dim ucConsultaProduto As ucConsultaProduto = Me.Page.FindControlRecursive("ucConsultaProduto")
            '            If ucConsultaProduto IsNot Nothing Then
            '                Session("Where" & HID.Value) = " Agrupar = 'S' AND NCM = '" & DsXml.Tables("prod").Rows(0)("NCM").ToString() & "'"
            '                ucConsultaProduto.Limpar()
            '                ucConsultaProduto.MainUserControl = Me
            '                ucConsultaProduto.SetarHID(HID.Value)
            '                ucConsultaProduto.BuscarProduto(True)
            '                Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            '                Popup.ConsultaDeProduto(Me.Page, "objProdutoNFGXml" & HID.Value, txtNome.ClientID, True)
            '            End If
            '        End If
            '    End If
            'End If

            Exit Sub
        Else

            objProdutoNota = objNotaFiscal.Itens(pPosicaoItem)
            If Not objProdutoNota.IUD = "I" Then objProdutoNota.IUD = "U"

            txtQuantidade.Text = objProdutoNota.QuantidadeFiscal
            txtUnitario.Text = objProdutoNota.Unitario
            txtValorTotal.Text = objProdutoNota.ValorTotal

            '#FimBaseDeCalculo
            'HIDBaseDeCalculo.Value = objProdutoNota.Produto.BaseCalculo
            SessaoSalvaProdutoNota()

        End If

        If Not objProdutoNota.Produto Is Nothing AndAlso objProdutoNota.Produto.CodigoGrupo.Length > 0 Then
            cmbGrupoProduto.SelectedValue = objProdutoNota.Produto.CodigoGrupo
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
            cmbProdutos.SelectedValue = objProdutoNota.CodigoProduto
        End If

        If objProdutoNota.CodigoOperacao <> 0 Then

            If Not itemRemovido Then txtOperacao.Text = objProdutoNota.CodigoOperacao & " - " &
           objProdutoNota.CodigoSubOperacao & "  " & objProdutoNota.SubOperacao.Descricao

            If objProdutoNota.Encargos.Count > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing Then
                divCentroDeCusto.Visible = objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
            End If


            If Not objProdutoNota.Encargos Is Nothing AndAlso
                objProdutoNota.Encargos.Count > 0 AndAlso
                Not objProdutoNota.Encargos.EncProduto Is Nothing AndAlso
                Not String.IsNullOrWhiteSpace(objProdutoNota.Encargos.EncProduto.CentroDeCusto) AndAlso
                CInt(objProdutoNota.Encargos.EncProduto.CentroDeCusto) > 0 Then
                cmbCentroCusto.SelectedValue = objProdutoNota.Encargos.EncProduto.CentroDeCusto
            End If
        End If



        'txtQuantidade.Text = objProdutoNota.QuantidadeFiscal.ToString("N4")
        'If objProdutoNota.CodigoProduto <> objNotaFiscal.Empresa.Empresa.CodigoProdutoDeFrete And
        '    objProdutoNota.CodigoProduto <> objNotaFiscal.Empresa.Empresa.CodigoProdutoDeEstadia And
        '    objProdutoNota.CodigoProduto <> objNotaFiscal.Empresa.Empresa.CodigoProdutoDeMDFe Then
        '    txtQuantidade.Enabled = True
        'End If

        'txtUnitario.Text = objProdutoNota.Unitario.ToString("N10")
        'txtValorTotal.Text = objProdutoNota.ValorTotal.ToString("N2")
        'txtValorTotalOld.Value = objProdutoNota.ValorTotal

        'txtUnitario.Text = dsXml.Tables("prod").Rows(pPosicaoItem).Item("vUnCom").ToString.Replace(".", ",")
        'txtValorTotal.Text = dsXml.Tables("prod").Rows(pPosicaoItem).Item("vProd").ToString.Replace(".", ",")
        'txtQuantidade.Text = dsXml.Tables("prod").Rows(pPosicaoItem).Item("qTrib").ToString.Replace(".", ",")
        txtValorTotalOld.Value = objProdutoNota.ValorTotal

        SetarMsgRetencao()
    End Sub

    Public Sub BloqueioEdicaoProdutoXML()
        lnkLimpar.Parent.Visible = False
        cmbGrupoProduto.Enabled = True
        cmbProdutos.Enabled = True
        BtnConsultaProduto.Visible = True
        imbOperacao.Visible = True
        txtUnitario.Enabled = False
        txtQuantidade.Enabled = False
        txtValorTotal.Enabled = False
    End Sub

    Public Sub BloqueioEdicaoProduto()
        lnkLimpar.Parent.Visible = False

        cmbGrupoProduto.Enabled = False
        cmbProdutos.Enabled = False
        BtnConsultaProduto.Visible = False

        'cmbGrupoProdutoCusto.Enabled = False
        'cmbProdutosCusto.Enabled = False
        'BtnConsultaProdutoCusto.Visible = False

        imbOperacao.Visible = False
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

#End Region

#Region "Variáveis Locais"
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objProdutoNota As [Lib].Negocio.NotaFiscalXItem
    Private objProdutoNotaOrigem As [Lib].Negocio.NotaFiscalXItem

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
            Dim objSubOperacao As [Lib].Negocio.SubOperacao = CType(obj, [Lib].Negocio.SubOperacao)
            Session.Remove("objOperacoesNFG" & HID.Value)

            Dim fCliente As String = objNotaFiscal.CodigoCliente
            Dim fEndCliente As Integer = objNotaFiscal.EnderecoCliente
            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then

                objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                    objNotaFiscal.NotasTrocaOrigem.Count > 0 Then
                    objNotaFiscal.CodigoCliente = objNotaFiscal.NotasTrocaOrigem(0).CodigoCliente
                    objNotaFiscal.EnderecoCliente = objNotaFiscal.NotasTrocaOrigem(0).EnderecoCliente
                End If
            End If

            If PosicaoItem.Value = -1 AndAlso objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count = 0 Then
                objNotaFiscal.CodigoOperacao = objSubOperacao.CodigoOperacao
                objNotaFiscal.CodigoSubOperacao = objSubOperacao.Codigo
            End If

            'If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count > 0 Then
            '    If PosicaoItem.Value = -1 Or PosicaoItem.Value > 0 Then
            '        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoOperacao <> objSubOperacao.CodigoOperacao Or objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoSubOperacao <> objSubOperacao.Codigo Then
            '            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).SubOperacao.Financeiro <> objSubOperacao.Financeiro Then
            '                MsgBox(Me.Page, "Produtos com operações diferentes devem estar configurados iguais no financeiro!")
            '                Exit Sub
            '            End If

            '            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Any(Function(s) s.SubOperacao.CodigoGrupoContas <> objSubOperacao.CodigoGrupoContas) Then
            '                MsgBox(Me.Page, "Produtos com operações diferentes devem ter o mesmo grupo de contas!")
            '                Exit Sub
            '            End If
            '        End If
            '    ElseIf objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count > 1 And PosicaoItem.Value = 0 Then
            '        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(1).CodigoOperacao <> objSubOperacao.CodigoOperacao Or objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(1).CodigoSubOperacao <> objSubOperacao.Codigo Then
            '            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(1).SubOperacao.Financeiro <> objSubOperacao.Financeiro Then
            '                MsgBox(Me.Page, "Produtos com operações diferentes devem estar configurados iguais no financeiro!")
            '                Exit Sub
            '            End If

            '            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Any(Function(s) s.SubOperacao.CodigoGrupoContas <> objSubOperacao.CodigoGrupoContas) Then
            '                MsgBox(Me.Page, "Produtos com operações diferentes devem ter o mesmo grupo de contas!")
            '                Exit Sub
            '            End If
            '        End If
            '    End If
            'End If

            txtOperacao.Text = String.Format("{0}-{1} - {2}", objSubOperacao.CodigoOperacao, objSubOperacao.Codigo, objSubOperacao.Descricao)
            SessaoRecuperaProdutoNota()

            Dim objProduto As New Produto(cmbProdutos.SelectedValue)
            Dim Parametros As New OperacaoXEstado

            Parametros.Empresa = Left(objNotaFiscal.Empresa.Codigo, 8)
            Parametros.CodigoGrupoProduto = objProduto.CodigoGrupo
            Parametros.CodigoProduto = objProduto.Codigo
            Parametros.CodigoOperacao = objSubOperacao.CodigoOperacao
            Parametros.CodigoSubOperacao = objSubOperacao.Codigo
            Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado
            Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado
            Parametros.Ativo = 1
            If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
            Parametros.InicioVigencia = objNotaFiscal.Movimento

            Dim itemCFOP As Integer = 0
            If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso (objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                itemCFOP = objSubOperacao.CodigoFiscal
            End If
            Dim OXE As New OperacaoXEstado(Parametros, True, itemCFOP)

            objProdutoNota.CodigoOperacao = objSubOperacao.CodigoOperacao
            objProdutoNota.CodigoSubOperacao = objSubOperacao.Codigo
            objNotaFiscal.CodigoOperacao = objSubOperacao.CodigoOperacao
            objNotaFiscal.CodigoSubOperacao = objSubOperacao.Codigo
            objProdutoNota.CodigoOperacaoEstado = OXE.Codigo
            objProdutoNota.CFOP = itemCFOP
            'objProdutoNota.CodigoFixacao

            'ACRESCENTADO PARA IR NO ESTOQUE FÍSICO CASO ESTEJA MARCADO NA OPERAÇÃO - FURLAN - 30/07/2024
            If objProdutoNota.SubOperacao.EstoqueFisico Then
                objProdutoNota.QuantidadeFisica = objProdutoNota.QuantidadeFiscal
            End If

            'Serve para Lista dos Encargos, gerar uma nova ao invés de tentar buscar o que está gravado na nota - Furlan - 25-01-2014
            'Na gravação CodigoTipoDeDocumentoFrete voltar a ser 0
            If Not objProdutoNota.Encargos Is Nothing AndAlso objProdutoNota.Encargos.Count = 0 Then objProdutoNota.Encargos = Nothing

            If objProdutoNota.Encargos Is Nothing OrElse objProdutoNota.Encargos.Count = 0 Then

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                    objNotaFiscal.TipoDeDocumentoFrete = Nothing

                    objNotaFiscal.CodigoCliente = fCliente
                    objNotaFiscal.EnderecoCliente = fEndCliente
                End If

                SessaoSalvaProdutoNota()
                SessaoSalvaNotaFiscal()

                MsgBox(Me.Page, "Não foram encontrados encargos do produto " & objProdutoNota.CodigoProduto & "-" & objProdutoNota.Produto.Descricao & " para a operação " & objSubOperacao.CodigoOperacao & "-" & objSubOperacao.Codigo)
                txtOperacao.Text = String.Empty
                Exit Sub
            End If

            If objProdutoNota.Encargos.Count > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing Then
                divProdutoCusto.Visible = objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.ProdutoParaCusto
            End If

            If Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto Then
                divCentroDeCusto.Visible = True
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
                divCentroDeCusto.Visible = False
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                'objNotaFiscal.TipoDeDocumentoFrete = Nothing

                objNotaFiscal.CodigoCliente = fCliente
                objNotaFiscal.EnderecoCliente = fEndCliente
            End If

            SessaoSalvaProdutoNota()
            SessaoSalvaNotaFiscal()
        ElseIf Session("objComissoesXBaixas" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim lstComissoesXBaixas As [Lib].Negocio.ListComissoesXBaixas = CType(Session("objComissoesXBaixas" & HID.Value), [Lib].Negocio.ListComissoesXBaixas)
            If Not (objNotaFiscal.ComissoesXBaixas IsNot Nothing AndAlso objNotaFiscal.ComissoesXBaixas.Count > 0) Then
                objNotaFiscal.ComissoesXBaixas = New ListComissoesXBaixas(objNotaFiscal)
            End If

            For Each objComissoesXBaixas As ComissoesXBaixas In lstComissoesXBaixas
                objNotaFiscal.ComissoesXBaixas.Add(objComissoesXBaixas)
            Next

            objNotaFiscal.ComissoesXBaixas.ForEach(Function(s)
                                                       If (Not s.Valor > 0) Then
                                                           s.IUD = "D"
                                                       Else
                                                           s.IUD = "I"
                                                       End If
                                                       Return True
                                                   End Function)
            SessaoSalvaNotaFiscal()
        End If
        SetarMsgRetencao()
    End Sub
#End Region

#Region "Botões"
    Protected Sub lnkGravar_Click(sender As Object, e As EventArgs) Handles lnkGravar.Click
        AdicionarProduto(False, ReaproveitarDados)
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click

        SessaoRecuperaProdutoNotaOrigem()

        'Se o usuario fechou sem fazer alteração, o item da nota não pode ser alterado e deve retornar aos valores original
        If Not objProdutoNotaOrigem Is Nothing Then

            SessaoRecuperaNotaFiscal()
            SessaoRecuperaProdutoNota()

            Dim index As Integer = 0
            For Each item As NotaFiscalXItem In objNotaFiscal.Itens()

                If item.Sequencia = objProdutoNotaOrigem.Sequencia Then
                    Exit For
                End If

                index += 1
            Next

            objProdutoNota = objProdutoNotaOrigem
            objNotaFiscal.Itens(index) = objProdutoNotaOrigem

            SessaoSalvaNotaFiscal()
            SessaoSalvaProdutoNota()

        End If

        Popup.CloseDialog(Me.Page, "divProdutoNFG")

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
              Or Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" _
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
                                                                          OrElse cmbGrupoProduto.SelectedValue = "50101" _
                                                                          OrElse cmbGrupoProduto.SelectedValue = "40201")) Then
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
        Else
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "' AND Agrupar = 'S' ", True)
        End If
    End Sub

    Protected Sub cmbProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProdutos.SelectedIndexChanged

        txtOperacao.Text = String.Empty

        SessaoRecuperaNotaFiscal()

        'Dim i As Integer = 0

        'If Not objNotaFiscal.Itens Is Nothing Then i = objNotaFiscal.Itens.FindIndex(Function(s) s.IUD <> "D" And s.CodigoProduto = cmbProdutos.SelectedValue)

        'If objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = cmbProdutos.SelectedValue AndAlso s.IUD = "D").Count > 0 Then
        '    Dim index As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = cmbProdutos.SelectedValue)
        '    InicializarUC(index)
        '    BloqueioEdicaoProduto()
        'Else -> une com o if abaixo elseif
        'If i >= 0 Then
        '    InicializarUC(i)
        'Else
        SessaoRecuperaProdutoNota()
        objProdutoNota.CodigoProduto = cmbProdutos.SelectedValue

        SessaoRecuperaProdutoNotaOrigem()

        If SessaoDsXML Is Nothing AndAlso objProdutoNotaOrigem Is Nothing Then
            objProdutoNota.Sequencia = objNotaFiscal.Itens.Count + 1
        Else
            objProdutoNota.Sequencia = objProdutoNota.Sequencia
        End If

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
        txtQuantidade.Text = String.Empty
        txtUnitario.Text = String.Empty
        txtValorTotal.Text = String.Empty
        txtValorTotalOld.Value = "0"
        cmbCentroCusto.SelectedIndex = 0
        divCentroDeCusto.Visible = False
        divProdutoCusto.Visible = False
        txtOperacao.Text = String.Empty
        PosicaoItem.Value = -1
        chkRetencao.Parent.Visible = False
        '#FimBaseDeCalculo
        'HIDBaseDeCalculo.Value = String.Empty
        If objNotaFiscal Is Nothing Then SessaoRecuperaNotaFiscal()
        objProdutoNota = New NotaFiscalXItem(objNotaFiscal)
        SessaoSalvaProdutoNota()

        objProdutoNotaOrigem = Nothing
        SessaoSalvaProdutoNotaOrigem()

        HttpContext.Current.Session("ssMessage") = String.Empty

    End Sub

    Public Sub AdicionarProduto(ByVal bXMLAutomatico As Boolean, ByVal bReaproveitarDados As Boolean)

        SessaoRecuperaNotaFiscal()
        SessaoRecuperaProdutoNota()

        Try

            If ValidarCamposProdutos(bXMLAutomatico) Then

                If SessaoDsXML Is Nothing Then
                    '#FimBaseDeCalculo
                    'txtValorTotal.Text = Math.Round((CDbl(txtQuantidade.Text) * CDbl(txtUnitario.Text)) / objProdutoNota.Produto.BaseCalculo, 2).ToString("N2")
                    txtValorTotal.Text = Math.Round((CDbl(txtQuantidade.Text) * CDbl(txtUnitario.Text)), 2).ToString("N2")
                End If

                'Serve para Lista dos Encargos, gerar uma nova ao invés de tentar buscar o que está gravado na nota - Furlan - 25-01-2014
                'Na gravação CodigoTipoDeDocumentoFrete voltar a ser 0

                Dim fCliente As String = objNotaFiscal.CodigoCliente
                Dim fEndCliente As Integer = objNotaFiscal.EnderecoCliente
                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                    objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                    If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                        objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                        objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                        objNotaFiscal.CodigoCliente = objNotaFiscal.NotasTrocaOrigem(0).CodigoCliente
                        objNotaFiscal.EnderecoCliente = objNotaFiscal.NotasTrocaOrigem(0).EnderecoCliente
                    End If
                End If

                If bXMLAutomatico = False Then
                    objProdutoNota.PesoQuantidade = objProdutoNota.Produto.PesoQuantidade
                    objProdutoNota.QuantidadeFiscal = txtQuantidade.Text

                    'ACRESCENTADO PARA IR NO ESTOQUE FÍSICO CASO ESTEJA MARCADO NA OPERAÇÃO - FURLAN - 30/07/2024
                    If Not objProdutoNota.SubOperacao.EstoqueFisico Then
                        objProdutoNota.QuantidadeFisica = 0
                    End If

                    If SessaoDsXML Is Nothing Then
                        objProdutoNota.Unitario = txtUnitario.Text
                        objProdutoNota.ValorTotal = CDec(txtValorTotal.Text)
                    End If

                End If

                'Kitio 18/07
                If PosicaoItem.Value >= 0 Then
                    For i As Integer = 0 To objNotaFiscal.Itens.Count - 1
                        If objNotaFiscal.Itens(i).IUD <> "D" Then
                            If PosicaoItem.Value = 0 Then
                                objNotaFiscal.CodigoOperacao = objProdutoNota.CodigoOperacao
                                objNotaFiscal.CodigoSubOperacao = objProdutoNota.CodigoSubOperacao
                            End If
                            'objNotaFiscal.Itens(i).Encargos = New ListNotaFiscalXItemXEncargo(objNotaFiscal.Itens(i))
                            If HttpContext.Current.Session("ssMessage") IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("ssMessage").ToString) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                HttpContext.Current.Session.Remove("ssMessage")
                            End If
                        End If
                    Next
                    objProdutoNota.Encargos = New ListNotaFiscalXItemXEncargo(objProdutoNota, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                ElseIf objProdutoNota.Encargos IsNot Nothing AndAlso objProdutoNota.Encargos.Count > 0 Then
                    Dim index As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = objProdutoNota.CodigoProduto And s.Sequencia = objProdutoNota.Sequencia And s.IUD = "D")
                    If index = -1 Then
                        If Not objProdutoNota.CodigoProduto Is Nothing Then

                            objProdutoNota.IUD = "I"
                            If bReaproveitarDados = False OrElse objNotaFiscal.Itens.Where(Function(x) x.CodigoProduto = objProdutoNota.CodigoProduto).Count = 0 Then
                                objNotaFiscal.Itens.Add(objProdutoNota)
                            End If

                            If objNotaFiscal.Itens.Where(Function(I) I.IUD <> "D").Count = 1 _
                            AndAlso objNotaFiscal.CodigoOperacao <> objProdutoNota.CodigoOperacao _
                            AndAlso objNotaFiscal.CodigoSubOperacao = objProdutoNota.CodigoSubOperacao Then
                                objNotaFiscal.CodigoOperacao = objProdutoNota.CodigoOperacao
                                objNotaFiscal.CodigoSubOperacao = objProdutoNota.CodigoSubOperacao
                            End If
                        End If
                    Else
                        objProdutoNota.IUD = "U"

                        objProdutoNota.CarregandoEncargos = True
                        objProdutoNota.Encargos = New ListNotaFiscalXItemXEncargo(objProdutoNota, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                        objProdutoNota.CarregandoEncargos = False

                        objProdutoNota.Sequencia = objNotaFiscal.Itens(index).Sequencia
                        objNotaFiscal.Itens(index) = objProdutoNota

                        If objNotaFiscal.Itens.Where(Function(I) I.IUD <> "D").Count = 1 _
                        AndAlso objNotaFiscal.CodigoOperacao <> objProdutoNota.CodigoOperacao _
                        AndAlso objNotaFiscal.CodigoSubOperacao = objProdutoNota.CodigoSubOperacao Then
                            objNotaFiscal.CodigoOperacao = objProdutoNota.CodigoOperacao
                            objNotaFiscal.CodigoSubOperacao = objProdutoNota.CodigoSubOperacao
                        End If
                    End If
                End If

                If TypeOf Me.Page Is NotasFiscaisGerais Then
                    CType(Me.Page, NotasFiscaisGerais).AtualizarProdutoXML()
                End If

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then

                    If objNotaFiscal.XMLImportado = False Then
                        objNotaFiscal.TipoDeDocumentoFrete = Nothing
                    End If

                    objNotaFiscal.CodigoCliente = fCliente
                    objNotaFiscal.EnderecoCliente = fEndCliente

                End If

                If objProdutoNota.Encargos IsNot Nothing AndAlso objProdutoNota.Encargos.Count > 0 Then
                    If cmbCentroCusto.SelectedIndex > 0 Then
                        objProdutoNota.Encargos.EncProduto.CentroDeCusto = cmbCentroCusto.SelectedValue
                        objProdutoNota.CentroDeCustoInformado = cmbCentroCusto.SelectedValue
                    End If
                End If

                If TypeOf Me.Page Is NotasFiscaisGerais AndAlso objProdutoNota.Encargos IsNot Nothing AndAlso objProdutoNota.Encargos.Count > 0 Then

                    'Repetir os dados para produtos similar nos outros itens produtos
                    For i As Integer = 0 To objNotaFiscal.Itens.Count - 1
                        Dim item As NotaFiscalXItem = objNotaFiscal.Itens(i)

                        ' Verifica a condição de filtro
                        If Not item.CodigoProduto Is Nothing AndAlso item.NCMProdutoXML = objProdutoNota.NCMProdutoXML AndAlso item.CodigoProduto.Length = 0 Then

                            ' Pula a iteração se for o mesmo índice
                            If objProdutoNota.Sequencia = item.Sequencia Then
                                Continue For
                            End If

                            ' Criando um clone do objProdutoNota
                            Dim itemClone As NotaFiscalXItem = objProdutoNota.Clone()

                            itemClone.Sequencia = item.Sequencia
                            ' Copiando os dados do item original para o clone
                            itemClone.PesoBruto = item.PesoBruto
                            itemClone.PesoFiscal = item.PesoFiscal
                            itemClone.PesoLiquido = item.PesoLiquido
                            itemClone.PesoQuantidade = item.PesoQuantidade
                            itemClone.QuantidadeDeEmbalagem = item.QuantidadeDeEmbalagem
                            itemClone.QuantidadeFiscal = item.QuantidadeFiscal
                            itemClone.QuantidadeFisica = item.QuantidadeFisica
                            itemClone.Unitario = item.Unitario
                            itemClone.ValorTotal = item.ValorTotal
                            itemClone.ValorTotalMoeda = item.ValorTotalMoeda
                            itemClone.ValorLiquido = item.ValorLiquido
                            itemClone.ValorLiquidoMoeda = item.ValorLiquidoMoeda

                            itemClone.DescricaoProdutoXML = item.DescricaoProdutoXML
                            itemClone.NomeProdutoXML = item.NomeProdutoXML
                            itemClone.NCMProdutoXML = item.NCMProdutoXML
                            itemClone.UnidadeProdutoXML = item.UnidadeProdutoXML
                            itemClone.ProdutoXML = item.ProdutoXML
                            itemClone.ValorDescontoXML = item.ValorDescontoXML
                            itemClone.ValorFreteXML = item.ValorFreteXML

                            itemClone.CarregandoEncargos = True
                            itemClone.Encargos = New ListNotaFiscalXItemXEncargo(itemClone, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                            itemClone.CarregandoEncargos = False

                            ' Atualizando a lista no índice correto
                            objNotaFiscal.Itens(i) = itemClone

                        End If
                    Next

                    SessaoSalvaProdutoNota()
                    SessaoSalvaNotaFiscal()

                    CType(Me.Page, NotasFiscaisGerais).AtualizarItensNoGrid()
                    CType(Me.Page, NotasFiscaisGerais).AtualizarValorParcelamento()

                    SessaoRecuperaProdutoNotaOrigem()

                    If Not objProdutoNotaOrigem Is Nothing Then
                        Popup.CloseDialog(Me.Page, "divProdutoNFG")
                    End If

                    Limpar()

                ElseIf TypeOf Me.Page Is NotasFiscaisGerais Then

                    If CType(Me.Page, NotasFiscaisGerais).VerificarProdutos() = False Then

                        If objProdutoNota.Encargos Is Nothing OrElse objProdutoNota.Encargos.Count() = 0 Then
                            MsgBox(Me.Page, String.Format("Encargos não encontrados para a operação: {0}-{1} - Origem: {2} - Destino: {3}!", objProdutoNota.CodigoOperacao, objProdutoNota.CodigoSubOperacao, objNotaFiscal.Empresa.CodigoEstado, objNotaFiscal.Cliente.CodigoEstado))
                            'Exit Sub
                        End If

                        Popup.ConsultaProdutoNFG(Me.Page, "objProdutoNotaNF" & HID.Value)
                    End If

                End If

            End If

        Catch ex As Exception

            MsgBox(Me.Page, ex.Message)

            Popup.CloseDialog(Me.Page, "divProdutoNFG")

            CType(Me.Page, NotasFiscaisGerais).LimparCampos()
        End Try

    End Sub

    Private Function ValidarCamposProdutos(ByVal bXMLAutomatico As Boolean) As Boolean

        If bXMLAutomatico Then
            Return True
        ElseIf cmbGrupoProduto.SelectedIndex < 1 Then
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
        ElseIf String.IsNullOrWhiteSpace(txtQuantidade.Text) Then
            MsgBox(Me.Page, "É necessário informar a quantidade!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtUnitario.Text) Then
            MsgBox(Me.Page, "É necessário informar o valor unitário do produto!")
            Return False
        ElseIf objProdutoNota.Encargos.Count > 0 AndAlso Not objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso objProdutoNota.Encargos.EncProduto.GrupoDeContaDebito.ProdutoParaCusto AndAlso (objProdutoNota.CodigoProdutoCusto Is Nothing OrElse objProdutoNota.CodigoProdutoCusto.Length = 0) Then
            MsgBox(Me.Page, "É necessário informar produto para custo!")
            Return False
        End If

        Return True

    End Function
#End Region

    Public Sub SetarMsgRetencao()
        '*************************************************************************************
        '*************** RETENCAO ************************************************************
        '*************************************************************************************
        SessaoRecuperaNotaFiscal()

        If PosicaoItem.Value = -1 OrElse objNotaFiscal.Itens(PosicaoItem.Value).OperacaoEstado Is Nothing Then
            If cmbProdutos.SelectedIndex = 0 Or (objProdutoNota Is Nothing OrElse objProdutoNota.CodigoSubOperacao = 0) Then
                chkRetencao.Parent.Visible = False
                chkRetencao.Checked = True

                If Not objProdutoNota Is Nothing Then
                    objProdutoNota.Retencao = True
                End If

                Exit Sub
            End If

            Dim objProduto As New Produto(cmbProdutos.SelectedValue)
            Dim Parametros As New OperacaoXEstado
            Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
            Parametros.CodigoGrupoProduto = objProduto.CodigoGrupo
            Parametros.CodigoProduto = objProduto.Codigo
            Parametros.CodigoOperacao = objProdutoNota.CodigoOperacao
            Parametros.CodigoSubOperacao = objProdutoNota.CodigoSubOperacao
            Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado
            Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado
            Parametros.Ativo = 1
            If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao

            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                    objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                    objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                    Parametros.EstadoDestino = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado
                    If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
                End If
            End If

            Parametros.InicioVigencia = objNotaFiscal.Movimento
            Dim OxE As New OperacaoXEstado(Parametros)

            If OxE.Encargos.Count = 0 Then
                chkRetencao.Parent.Visible = False
                chkRetencao.Checked = True
                objProdutoNota.Retencao = True
                Exit Sub
            End If

            objProdutoNota.CodigoOperacaoEstado = OxE.Codigo

            If OxE.CodigoNaturezaDeRendimento > 0 Then
                objNotaFiscal.CodigoNaturezaDeRendimento = OxE.CodigoNaturezaDeRendimento
                SessaoSalvaNotaFiscal()
            End If

            If objNotaFiscal.Cliente.CodigoEstado <> "EX" Then
                If objNotaFiscal.Cliente.Codigo.Length = 14 Then
                    If OxE.Encargos.DescRetencaoPJ.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & OxE.Encargos.DescRetencaoPJ
                        chkRetencao.Checked = True
                        objProdutoNota.Retencao = True
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = True
                        objProdutoNota.Retencao = True
                    End If
                Else
                    If OxE.Encargos.DescRetencaoPF.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & OxE.Encargos.DescRetencaoPF
                        chkRetencao.Checked = True
                        objProdutoNota.Retencao = True
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = True
                        objProdutoNota.Retencao = True
                    End If
                End If
            End If
        Else
            If objNotaFiscal.Cliente.CodigoEstado <> "EX" Then
                If objNotaFiscal.Cliente.Codigo.Length = 14 Then
                    If objNotaFiscal.Itens(PosicaoItem.Value).OperacaoEstado.Encargos.DescRetencaoPJ.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & objNotaFiscal.Itens(PosicaoItem.Value).OperacaoEstado.Encargos.DescRetencaoPJ
                        chkRetencao.Checked = objNotaFiscal.Itens(PosicaoItem.Value).Retencao
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = True
                        objProdutoNota.Retencao = True
                    End If
                Else
                    If objNotaFiscal.Itens(PosicaoItem.Value).OperacaoEstado.Encargos.DescRetencaoPF.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & objNotaFiscal.Itens(PosicaoItem.Value).OperacaoEstado.Encargos.DescRetencaoPF
                        chkRetencao.Checked = objNotaFiscal.Itens(PosicaoItem.Value).Retencao
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = True
                        objProdutoNota.Retencao = True
                    End If
                End If
            End If
        End If
        SessaoSalvaProdutoNota()
    End Sub

    Protected Sub chkRetencao_CheckedChanged(sender As Object, e As EventArgs) Handles chkRetencao.CheckedChanged
        SessaoRecuperaProdutoNota()
        objProdutoNota.Retencao = chkRetencao.Checked
        SessaoSalvaProdutoNota()
    End Sub
End Class