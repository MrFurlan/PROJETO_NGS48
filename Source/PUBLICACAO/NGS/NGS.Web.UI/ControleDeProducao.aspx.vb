Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ControleDeProducao
    Inherits BasePage

#Region "Variáveis"

    Dim blnModoInclusao As Boolean = True
    Dim Carrega As New CarregarDDL
    Dim Prod As [Lib].Negocio.Produto

#End Region

#Region "Tab Manutenção"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ControleDeProducao", "ACESSAR") Then
                CarregarUnidades()
                'VerificaUnidade()
                CarregarEtapas()
                CarregarGruposProdutos()
                'txtMovimento.Text = DateTime.Now.ToString("dd/MM/yyyy")
                'HID.Value = Guid.NewGuid().ToString
                'ucConsultaClientes.SetarHID(HID.Value)
                'ucConsultaOperacoes.SetarHID(HID.Value)
                LimparCampos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteCTRPROD" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteCTRPROD" & HID.Value), [Lib].Negocio.Cliente))
            txtDeposito.Text = itemCliente.Text
            txtCodigoDeposito.Value = itemCliente.Value
            Session.Remove("objClienteCTRPROD" & HID.Value)
        ElseIf Session("objSubOperacao" & HID.Value) IsNot Nothing Then
            Dim objSubOperacao As [Lib].Negocio.SubOperacao = CType(obj, [Lib].Negocio.SubOperacao)

            If objSubOperacao.EstoqueFisico Then
                chkFisico.Checked = True
            Else
                chkFisico.Checked = False
            End If

            If objSubOperacao.EstoqueFiscal Then
                chkFiscal.Checked = True
            Else
                chkFiscal.Checked = False
            End If

            txtOperacao.Text = String.Format("{0}-{1} - {2}", objSubOperacao.CodigoOperacao, objSubOperacao.Codigo, objSubOperacao.Descricao)
            txtCodigoOperacao.Value = String.Format("{0};{1};{2}", objSubOperacao.CodigoOperacao, objSubOperacao.Codigo, IIf(objSubOperacao.EntradaSaida = eEntradaSaida.Entrada, "E", "S"))

            If Not objSubOperacao.GrupoDeConta Is Nothing AndAlso objSubOperacao.GrupoDeConta.TemCentroDeCusto Then
                divCentroDeCusto.Visible = True
                CarregarCentrodeCusto()
            Else
                divCentroDeCusto.Visible = False
                ddlCentroDeCusto.DataSource = Nothing
            End If

            Session.Remove("objSubOperacao" & HID.Value)
        ElseIf Session("objProdutoETQ" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)
            cmbGruposProdutos.SelectedValue = objProduto.CodigoGrupo

            cmbProdutos.Items.Clear()
            ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, " Situacao = 1")
            cmbProdutos.SelectedValue = objProduto.Codigo.Trim()

            If blnModoInclusao Then If cmbProdutos.SelectedIndex > 0 Then CarregarAnalisesInclusao()

            Prod = New [Lib].Negocio.Produto(cmbProdutos.SelectedValue)
            LblUn.Text = Prod.Unidade

            If Prod.ControlarLote Then
                ddlLote.Enabled = True
                ddlLoteClassificacao.Enabled = True
                Carrega.Carregar(ddlLote, CarregarDDL.Tabela.LoteProduto, "Produto_id = '" & Prod.Codigo & "' and DataValidade >= '" & txtMovimento.Text.ToSqlDate() & "'")
            Else
                ddlLote.Enabled = False
                ddlLoteClassificacao.Enabled = False
                ddlLote.Items.Clear()
                ddlLoteClassificacao.Items.Clear()
            End If

            If Prod.ControlarNumeroDoLote Then
                idControlarNumeroDoLote.Visible = True
                imgSelecionaLote.Visible = True
                txtNumeroDoLote.Text = String.Empty

                Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")

                If strOperacao(2) = "E" Then
                    txtNumeroDoLote.Enabled = True
                    imgSelecionaLote.Visible = False
                    txtDataFabricacao.Enabled = True
                    txtDataValidade.Enabled = True
                Else
                    txtDataFabricacao.Enabled = False
                    txtDataValidade.Enabled = False
                End If
            Else
                txtNumeroDoLote.Text = String.Empty
                imgSelecionaLote.Visible = False
                idControlarNumeroDoLote.Visible = False
            End If

            If Prod.ControlarEmbalagem Then
                ddlEmbalagem.Enabled = True
                Carrega.Carregar(ddlEmbalagem, CarregarDDL.Tabela.EmbalagemTipoQtdeDoProduto, Prod.Codigo)
            Else
                ddlEmbalagem.Enabled = False
                ddlEmbalagem.Items.Clear()
            End If

            Session.Remove("objProdutoETQ" & HID.Value)

        ElseIf Session("objLoteFornecedor" & HID.Value) IsNot Nothing Then
            For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                If drItemLote("Produto") = cmbProdutos.SelectedValue Then
                    If drItemLote("Consumo") > 0 Then
                        txtNumeroDoLote.Text = drItemLote("Lote")
                        txtDataFabricacao.Text = drItemLote("Fabricado")
                        txtDataValidade.Text = drItemLote("Validade")
                        Exit Sub
                    End If
                End If
            Next

        End If
    End Sub

    Protected Sub cmbUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnidade.SelectedIndexChanged
        If cmbUnidade.SelectedIndex > 0 Then
            CarregarEmpresas()
        End If
    End Sub

    Protected Sub cmdConsultaDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        HttpContext.Current.Session("ssCampo") = "Livre"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCTRPROD" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmdOperacao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If String.IsNullOrWhiteSpace(cmbEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar a empresa!")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(txtCodigoDeposito.Value) Then
            MsgBox(Me.Page, "É necessário selecionar o depósito!")
            Exit Sub
        End If

        HttpContext.Current.Session("ssCampo") = "Livre"
        Dim strEmpresa = cmbEmpresa.SelectedValue.Split("-")
        Dim objEmpresa = New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
        Dim strCliente = txtCodigoDeposito.Value.Split("-")
        Dim objCliente = New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
        Dim parameters As New Dictionary(Of String, Object)
        parameters("tipo") = "Producao"
        parameters("documento") = ""
        parameters("prod") = cmbProdutos.SelectedValue
        parameters("ufOri") = IIf(objEmpresa Is Nothing, "", objEmpresa.CodigoEstado)
        parameters("cliDes") = IIf(objCliente Is Nothing, "", objCliente.Codigo & "-" & objCliente.CodigoEndereco)
        Popup.ConsultaDeOperacoes(Me.Page, "objSubOperacao" & HID.Value)
        ucConsultaOperacoes.Limpar()
        ucConsultaOperacoes.BindGridView(parameters)
    End Sub

    Protected Sub cmbGruposProdutosDerivados_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGruposProdutosDerivados.SelectedIndexChanged
        If cmbGruposProdutosDerivados.SelectedIndex > 0 Then
            CarregarProdutos("Derivados")
        End If
    End Sub

    Protected Sub cmbGruposProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGruposProdutos.SelectedIndexChanged

        If txtCodigoOperacao.Value.Length = 0 Then
            MsgBox(Me.Page, "Operação deve ser selecionada antes.")
            cmbGruposProdutos.SelectedIndex = 0
            Exit Sub
        End If

        If cmbGruposProdutos.SelectedIndex > 0 Then
            CarregarProdutos("Normal")
        End If
    End Sub

    Protected Sub cmbProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProdutos.SelectedIndexChanged
        If cmbProdutos.SelectedIndex > 0 Then
            cmbGruposProdutosDerivados.Focus()

            If blnModoInclusao Then If cmbProdutos.SelectedIndex > 0 Then CarregarAnalisesInclusao()

            Prod = New [Lib].Negocio.Produto(cmbProdutos.SelectedValue)
            LblUn.Text = Prod.Unidade

            If Prod.ControlarLote Then
                ddlLote.Enabled = True
                ddlLoteClassificacao.Enabled = True
                Carrega.Carregar(ddlLote, CarregarDDL.Tabela.LoteProduto, "Produto_id = '" & Prod.Codigo & "' and DataValidade >= '" & txtMovimento.Text.ToSqlDate() & "'")
            Else
                ddlLote.Enabled = False
                ddlLoteClassificacao.Enabled = False
                ddlLote.Items.Clear()
                ddlLoteClassificacao.Items.Clear()
            End If

            If Prod.ControlarNumeroDoLote Then
                idControlarNumeroDoLote.Visible = True
                imgSelecionaLote.Visible = True
                txtNumeroDoLote.Text = String.Empty

                Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")

                If strOperacao(2) = "E" Then
                    txtNumeroDoLote.Enabled = True
                    imgSelecionaLote.Visible = False
                    txtDataFabricacao.Enabled = True
                    txtDataValidade.Enabled = True
                Else
                    txtDataFabricacao.Enabled = False
                    txtDataValidade.Enabled = False
                End If

            Else
                txtNumeroDoLote.Text = String.Empty
                imgSelecionaLote.Visible = False
                idControlarNumeroDoLote.Visible = False
            End If

            If Prod.ControlarEmbalagem Then
                ddlEmbalagem.Enabled = True
                Carrega.Carregar(ddlEmbalagem, CarregarDDL.Tabela.EmbalagemTipoQtdeDoProduto, Prod.Codigo)
            Else
                ddlEmbalagem.Enabled = False
                ddlEmbalagem.Items.Clear()
            End If
        End If
    End Sub

    Protected Sub cmbEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbEmpresa.SelectedIndexChanged
        If cmbEmpresa.SelectedIndex > 0 Then
            txtDeposito.Text = cmbEmpresa.SelectedItem.Text
            txtCodigoDeposito.Value = cmbEmpresa.SelectedValue
        Else
            txtDeposito.Text = ""
            txtCodigoDeposito.Value = ""
        End If
    End Sub

    'Protected Sub TipoEstoque_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If Not chkFisico.Checked And Not chkFiscal.Checked Then
    '        chkFisico.Checked = (CType(sender, CheckBox).ID = "chkFiscal")
    '        chkFiscal.Checked = (CType(sender, CheckBox).ID = "chkFisico")
    '    End If
    'End Sub

#End Region

#Region "Procedimentos"

    Private Sub CarregarUnidades()
        Carrega.Carregar(cmbUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(cmbUnidade, cmbEmpresa)
        txtCodigoDeposito.Value = cmbEmpresa.SelectedValue
        txtDeposito.Text = cmbEmpresa.SelectedItem.Text
    End Sub

    Private Sub CarregarEmpresas()
        Carrega.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidade.SelectedValue, True)
    End Sub

    Private Sub CarregarEtapas()
        Carrega.Carregar(cmbEtapas, CarregarDDL.Tabela.Etapas, "", True)

        If cmbEtapas.Items.Count > 1 Then
            cmbEtapas.SelectedIndex = -1
            Dim itemEtapa As ListItem = cmbEtapas.Items.FindByValue("1")
            itemEtapa.Selected = True
        Else : cmbEtapas.SelectedIndex = 0
        End If
    End Sub

    Private Sub CarregarGruposProdutos()
        Carrega.Carregar(cmbGruposProdutos, CarregarDDL.Tabela.GrupoProduto, "", True)
        Carrega.Carregar(cmbGruposProdutosDerivados, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub LimparCampos()
        Session.Remove("objLoteFornecedor" & HID.Value)
        Session.Remove("objClienteCTRPROD" & HID.Value)
        Session.Remove("objSubOperacao" & HID.Value)
        Session.Remove("objProdutoETQ" & HID.Value)
        Session.Remove("dsAnalises" & HID.Value)
        Session.Remove("Where" & HID.Value)
        Session.Remove("dsProducao" & HID.Value)

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        idAnalises.Visible = False

        idControlarNumeroDoLote.Visible = False
        imgSelecionaLote.Visible = False
        txtNumeroDoLote.Text = String.Empty

        VerificaUnidade()
        txtMovimento.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtMovimentoHora.Text = "00:00"
        chkFisico.Checked = False
        chkFiscal.Checked = False

        cmbEtapas.SelectedIndex = -1
        Dim itemEtapa As ListItem = cmbEtapas.Items.FindByValue("1")
        itemEtapa.Selected = True

        txtOperacao.Text = ""
        txtCodigoOperacao.Value = ""
        cmbGruposProdutos.SelectedIndex = 0
        cmbProdutos.Items.Clear()
        cmbGruposProdutosDerivados.SelectedIndex = 0
        cmbProdutosDerivados.Items.Clear()
        ddlLote.Items.Clear()
        ddlLoteClassificacao.Items.Clear()
        ddlEmbalagem.Items.Clear()
        txtQuantidade.Text = ""
        txtObservacao.Text = ""
        LblUn.Text = ""

        txtNumeroDoLote.Enabled = False

        grdConsulta.DataSource = Nothing
        grdConsulta.DataBind()

        gridAnalises.DataSource = Nothing
        gridAnalises.DataBind()

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaOperacoes.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ucConsultaLote.SetarHID(HID.Value)

        Dim dtLoteFornecedor As New DataTable("ItemLoteFornecedor")
        dtLoteFornecedor.Columns.Add("Produto", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Lote", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Fabricado", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Validade", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtLoteFornecedor.Columns.Add("Consumo", Type.GetType("System.Decimal"))
        Session("objLoteFornecedor" & HID.Value) = dtLoteFornecedor

        HabilitarCampos(True)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbUnidade.Enabled = False
            cmbEmpresa.Enabled = False
        End If
    End Sub

    'Private Sub LimparCampos(ByVal Parcial As Boolean)
    '    txtOperacao.Text = ""
    '    txtCodigoOperacao.Value = ""
    '    cmbGruposProdutos.SelectedIndex = 0
    '    cmbProdutos.Items.Clear()
    '    cmbGruposProdutosDerivados.SelectedIndex = 0
    '    cmbProdutosDerivados.Items.Clear()
    '    ddlLote.Items.Clear()
    '    ddlLoteClassificacao.Items.Clear()
    '    ddlEmbalagem.Items.Clear()
    '    txtQuantidade.Text = ""
    '    txtObservacao.Text = ""

    '    chkFiscal.Checked = False
    '    chkFisico.Checked = False

    '    Session.Remove("dsProducao")
    '    grdConsulta.DataSource = Nothing
    '    grdConsulta.DataBind()

    '    gridAnalises.DataSource = Nothing
    '    gridAnalises.DataBind()

    '    HabilitarCampos(True)
    'End Sub

    Public Sub ConsultarProducao()
        Dim strSQL As String = "SELECT GE.Empresa_Id AS GrupoEmpresa, P.Empresa_Id AS CNPJEmpresa, P.EndEmpresa_Id, " & vbCrLf &
                               "       E.Reduzido AS ReduzidoEmpresa, P.Deposito_Id AS CNPJDeposito, P.EndDeposito_Id, " & vbCrLf &
                               "       D.Nome AS NomeDeposito, D.Reduzido AS ReduzidoDeposito, D.Cidade AS CidadeDeposito, " & vbCrLf &
                               "       D.Estado AS EstadoDeposito, P.Movimento_Id AS Movimento, " & vbCrLf &
                               "       P.Lote_Id, P.Classificacao_Id, P.Observacao, P.CentroDeCusto, " & vbCrLf &
                               "       case when P.Embalagem_Id = 0 then '' else convert(nvarchar,isnull(P.Embalagem_Id,'')) + ';' +  isnull(P.TipoDeEmbalagem_Id,'') +';'+ convert(nvarchar, P.CapacidadeEmbalagem_Id) end as DDLCodigoEmbalagem," & vbCrLf &
                               "       case when P.Embalagem_Id = 0 then '' else convert(nvarchar,isnull(Emb.EmbalagemIndea,'')) + '-'+ isnull(Emb.Descricao,'') +' / ' +  isnull(P.TipoDeEmbalagem_Id,'') +'-'+ isnull(TE.Descricao,'')+ '.....:'+ convert(nvarchar, P.CapacidadeEmbalagem_Id) end as DDLDescricaoEmbalagem," & vbCrLf &
                               "       convert(nvarchar,isnull(Emb.EmbalagemIndea,'')) + '-' +  isnull(P.TipoDeEmbalagem_Id,'') +' / '+ convert(nvarchar, P.CapacidadeEmbalagem_Id) as Embalagem," & vbCrLf
        '"       P.Embalagem_Id, P.TipoDeEmbalagem_Id, P.QtdeDeEmbalagem_Id as Embalagem" & vbCrLf & _
        strSQL &= "       CASE " & vbCrLf &
                               "         WHEN P.FisicoFiscal_Id = 1 " & vbCrLf &
                               "           THEN 'Fisico'" & vbCrLf &
                               "           ELSE 'Fiscal'" & vbCrLf &
                               "       END AS Tipo," & vbCrLf &
                               "       P.Etapa, P.Operacao_Id, " & vbCrLf &
                               "       P.SubOperacao_Id, SO.Descricao AS NomeOperacao, SO.EntradaSaida," & vbCrLf &
                               "       SO.Devolucao + ';' + SO.PrecoFixo + ';' + SO.Laudo + ';' + SO.EstoqueInicial + ';' + SO.EstoqueFisico + ';' + SO.EstoqueFiscal + ';' + SO.QuantidadeFisico + ';' + SO.QuantidadeFiscal + ';' + SO.QuantidadePedido + ';' + SO.UnitarioPedido + ';' + SO.Financeiro + ';' + SO.Contabil AS Flags," & vbCrLf &
                               "       PR.Grupo, P.Produto_Id, " & vbCrLf &
                               "       P.Produto_Id + '-' + PR.Nome AS ValorProduto, PD.Grupo AS GrupoDerivado, " & vbCrLf &
                               "       P.ProdutoDerivado_Id, " & vbCrLf &
                               "       CASE " & vbCrLf &
                               "         WHEN PD.Nome IS NOT NULL " & vbCrLf &
                               "           THEN P.ProdutoDerivado_Id + '-' + PD.Nome " & vbCrLf &
                               "           ELSE '' " & vbCrLf &
                               "       END AS ValorProdutoDerivado, " & vbCrLf &
                               "       CASE " & vbCrLf &
                               "         WHEN P.Entradas = 0 " & vbCrLf &
                               "           THEN P.Saidas " & vbCrLf &
                               "           ELSE P.Entradas" & vbCrLf &
                               "       END AS Quantidade, " & vbCrLf &
                               "       REPLICATE('0', 2 - LEN(CAST(P.Operacao_Id AS VARCHAR))) + CAST(P.Operacao_Id AS VARCHAR)+ '-' + REPLICATE('0', 3 - LEN(CAST(P.SubOperacao_Id AS VARCHAR))) + CAST(P.SubOperacao_Id AS VARCHAR) AS Operacao, ISNULL(OrdemDeProducao,0) AS OrdemDeProducao, " & vbCrLf &
                               "       ISNULL(Fabricado,'') AS Fabricado, ISNULL(Validade,'') AS Validade " & vbCrLf &
                               "  FROM Producao P " & vbCrLf &
                               " INNER JOIN (SELECT GE.Cliente_Id, GE.EndCliente_Id, MIN(GE.Empresa_Id) AS Empresa_Id " & vbCrLf &
                               "               FROM GruposXEmpresas GE " & vbCrLf &
                               "              INNER JOIN ClientesXTipos CT " & vbCrLf &
                               "                 ON GE.Empresa_Id = CT.Cliente_Id " & vbCrLf &
                               "                AND GE.EndEmpresa_Id = CT.Endereco_Id " & vbCrLf &
                               "              WHERE CT.Tipo_Id = 50 " & vbCrLf &
                               "              GROUP BY GE.Cliente_Id, GE.EndCliente_Id) GE " & vbCrLf &
                               "    ON GE.Cliente_Id = P.Empresa_Id " & vbCrLf &
                               "   AND GE.EndCliente_Id = P.EndEmpresa_Id " & vbCrLf &
                               " INNER JOIN Clientes E " & vbCrLf &
                               "    ON E.Cliente_Id = P.Empresa_Id " & vbCrLf &
                               "   AND E.Endereco_Id = P.EndEmpresa_Id " & vbCrLf &
                               " INNER JOIN Clientes D " & vbCrLf &
                               "    ON D.Cliente_Id = P.Deposito_Id " & vbCrLf &
                               "   AND D.Endereco_Id = P.EndDeposito_Id " & vbCrLf &
                               " INNER JOIN SubOperacoes SO " & vbCrLf &
                               "    ON SO.Operacao_Id = P.Operacao_Id " & vbCrLf &
                               "   AND SO.SubOperacoes_Id = P.SubOperacao_Id " & vbCrLf &
                               " INNER JOIN Produtos PR " & vbCrLf &
                               "    ON PR.Produto_Id = P.Produto_Id " & vbCrLf &
                               "  LEFT JOIN Produtos PD " & vbCrLf &
                               "    ON PD.Produto_Id = P.ProdutoDerivado_Id " & vbCrLf &
                               "  Left Join Embalagens Emb " & vbCrLf &
                               "    on Emb.Embalagem_id = P.Embalagem_Id " & vbCrLf &
                               "  Left Join TipoDeEmbalagem TE " & vbCrLf &
                               "    on TE.TipoDeEmbalagem_id = P.TipoDeEmbalagem_id " & vbCrLf &
                               " WHERE EXISTS (SELECT NULL " & vbCrLf &
                                                "FROM Clientes C " & vbCrLf &
                                               "WHERE C.Cliente_Id = GE.Empresa_Id) "

        If chkFisico.Checked And Not chkFiscal.Checked Then
            strSQL &= "AND P.FisicoFiscal_Id = 1 " & vbCrLf
        ElseIf chkFiscal.Checked And Not chkFisico.Checked Then
            strSQL &= "AND P.FisicoFiscal_Id = 2 " & vbCrLf
        End If

        If cmbEmpresa.SelectedIndex > 0 Then
            Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
            strSQL &= "AND P.Empresa_Id = '" & strEmpresa(0) & "' " & vbCrLf &
                      "AND P.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
        End If

        If txtCodigoDeposito.Value.Length > 0 Then
            Dim strDeposito As String() = txtCodigoDeposito.Value.Split(";")
            If strDeposito.Length > 1 = False Then strDeposito = txtCodigoDeposito.Value.Split("-")
            strSQL &= "AND P.Deposito_Id = '" & strDeposito(0) & "' " & vbCrLf &
                      "AND P.EndDeposito_Id = " & strDeposito(1) & " " & vbCrLf
        End If

        If txtMovimento.Text.Replace("/", "").Replace("_", "").Length > 0 Then
            strSQL &= "AND P.Movimento_Id BETWEEN '" & txtMovimento.Text.ToSqlDate() & "' " & vbCrLf &
                      "AND '" & Convert.ToDateTime(txtMovimento.Text).AddSeconds(86399).ToString("yyyy-MM-dd HH:mm:ss") & "' " & vbCrLf
        End If

        If cmbProdutos.SelectedIndex > 0 Then
            strSQL &= "AND P.Produto_Id = '" & cmbProdutos.SelectedValue & "' " & vbCrLf
        ElseIf cmbGruposProdutos.SelectedIndex > 0 Then
            strSQL &= "AND P.Produto_Id LIKE '" & cmbGruposProdutos.SelectedValue & "%' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoOperacao.Value) Then
            Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")
            strSQL &= "AND P.Operacao_Id = " & strOperacao(0) & vbCrLf &
                      "AND P.SubOperacao_Id = " & strOperacao(1) & vbCrLf
        End If

        Dim dsProducao As DataSet = Banco.ConsultaDataSet(strSQL, "Producao")
        grdConsulta.DataSource = dsProducao
        grdConsulta.DataBind()

        Session("dsProducao" & HID.Value) = dsProducao
    End Sub

    Private Function ValidarCampos() As Boolean
        Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")
        Dim objSubOperacao As [Lib].Negocio.SubOperacao = New [Lib].Negocio.SubOperacao(strOperacao(0), strOperacao(1))

        If cmbProdutos.SelectedIndex < 0 Then
            MsgBox(Me.Page, "Informe o produto para este registro.")
            Return False
        End If

        If Not IsNumeric(txtQuantidade.Text) OrElse (CDec(txtQuantidade.Text) <= 0) Then
            MsgBox(Me.Page, "Informe a Quantidade do Produto.")
            Return False
        End If

        Dim Prod As New [Lib].Negocio.Produto(cmbProdutos.SelectedValue)
        If Prod.ControlarLote Then
            If ddlLote.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Informe o Lote do Produto.")
                Return False
            ElseIf ddlLoteClassificacao.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Informe a Classificacao do Lote do Produto.")
                Return False
            End If
        End If

        If Prod.ControlarNumeroDoLote Then
            If (String.IsNullOrWhiteSpace(txtNumeroDoLote.Text) Or String.IsNullOrEmpty(txtNumeroDoLote.Text)) Then
                MsgBox(Me.Page, "Informe o número do Lote do Produto.")
                Return False
            ElseIf Not IsDate(txtDataFabricacao.Text) Then
                MsgBox(Me.Page, "Informe uma data de Fabricação válida.")
                Return False
            ElseIf Not IsDate(txtDataValidade.Text) Then
                MsgBox(Me.Page, "Informe uma data de Validade válida.")
                Return False
            ElseIf Not Left(cmbEmpresa.SelectedValue, 8) = "05272759" Then
                If objSubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso CDate(txtDataFabricacao.Text) < CDate(txtMovimento.Text) Then
                    MsgBox(Me.Page, "Data de Fabricação não pode ser menor que a Data do Movimento.")
                    Return False
                End If
            ElseIf CDate(txtDataValidade.Text) < CDate(txtDataFabricacao.Text) Then
                MsgBox(Me.Page, "Data de Validade não pode ser menor que a Data de Fabricação.")
                Return False
            End If
        End If

        If Prod.ControlarEmbalagem Then
            If ddlEmbalagem.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Informe a Embalagem do Produto.")
                Return False
            End If
            If Prod.ControlarLote Then
                Dim Emb As String() = ddlEmbalagem.SelectedValue.Split(";")
                Dim EmbalagemProduto As New [Lib].Negocio.ProdutoXEmbalagem(Prod.Codigo, Emb(0), Emb(1), Emb(2).Replace(".", ","))
                If EmbalagemProduto.PesoVariavel Then
                    Dim PeneiraDoLote As New [Lib].Negocio.LoteXClassificacao(Prod.Codigo, ddlLote.SelectedValue, ddlLoteClassificacao.SelectedValue)
                    If (CDec(txtQuantidade.Text) Mod PeneiraDoLote.PesoSaco) > 0 Then
                        MsgBox(Me.Page, "O Peso do Saco desta peneira é de " & Str(PeneiraDoLote.PesoSaco) & " / a Quantidade do produto de ser multiplo deste peso.")
                        Return False
                    End If
                End If
            End If
        End If

        If cmbEmpresa.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Informe os dados da empresa.")
            Return False
        ElseIf txtCodigoDeposito.Value.Length = 0 Then
            MsgBox(Me.Page, "Informe os dados do depósito.")
            Return False
        ElseIf txtCodigoOperacao.Value.Length = 0 Then
            MsgBox(Me.Page, "Informe a operação para este registro.")
            Return False
        ElseIf txtMovimento.Text.Replace("/", "").Replace("_", "").Length <= 0 Then
            MsgBox(Me.Page, "Informe a data de movimento deste registro.")
            Return False
        Else
            If Not objSubOperacao.EstoqueFisico And chkFisico.Checked Then
                MsgBox(Me.Page, "Esta operação não está liberada para receber lançamentos no estoque físico.")
                Return False
            ElseIf Not objSubOperacao.EstoqueFiscal And chkFiscal.Checked Then
                MsgBox(Me.Page, "Esta operação não está liberada para receber lançamentos no estoque fiscal.")
                Return False
            End If
        End If

        If Not Funcoes.VerificaAcesso(cmbEmpresa.SelectedValue.Split("-")(0), cmbEmpresa.SelectedValue.Split("-")(1), txtMovimento.Text, "PRODUCAO") Then
            MsgBox(Me.Page, "Movimento de Produção já Fechado para esta data...")
            Return False
        End If

        If Not chkFisico.Checked And Not chkFiscal.Checked Then
            MsgBox(Me.Page, "Marque o Tipo de Estoque para o lançamento.")
            Return False
        End If

        If ddlCentroDeCusto.Visible = True AndAlso (ddlCentroDeCusto.SelectedValue.ToString().Length = 0 OrElse ddlCentroDeCusto.SelectedValue = -1 OrElse ddlCentroDeCusto.SelectedValue = 0) Then
            MsgBox(Me.Page, "Informe o centro de custo.")
            Return False
        End If

        Return True
    End Function

    Private Sub CarregarProdutos(ByVal Tipo As String)
        Dim cmbCombo As DropDownList = IIf(Tipo = "Normal", cmbProdutos, cmbProdutosDerivados)
        Dim cmbGrupoCombo As DropDownList = IIf(Tipo = "Normal", cmbGruposProdutos, cmbGruposProdutosDerivados)

        cmbCombo.Items.Clear()

        Carrega.Carregar(cmbCombo, CarregarDDL.Tabela.Produto, "Grupo = " & cmbGrupoCombo.SelectedValue, True)
    End Sub

    Private Sub CarregarAnalisesInclusao()
        Dim strSQL As String = "SELECT PA.Analise_Id, A.Descricao, A.IndiceMinimo, A.IndiceMaximo, NULL AS Indice " &
                               "FROM ProdutosXAnalises PA " &
                               "INNER JOIN Analises A " &
                               "ON A.Analise_Id = PA.Analise_Id " &
                               "WHERE PA.Produto_Id = '" & cmbProdutos.SelectedValue & "' " &
                               "ORDER BY PA.Analise_Id"

        Dim dsAnalises As DataSet = Banco.ConsultaDataSet(strSQL, "ProdutosXAnalises")

        If dsAnalises.Tables(0).Rows().Count > 0 Then
            idAnalises.Visible = True
        Else
            idAnalises.Visible = False
        End If

        gridAnalises.DataSource = dsAnalises
        gridAnalises.DataBind()

        Session("dsAnalises" & HID.Value) = dsAnalises
    End Sub

    Protected Sub ddlLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlLote.SelectedIndexChanged
        If ddlLote.SelectedIndex > 0 Then
            Carrega.Carregar(ddlLoteClassificacao, CarregarDDL.Tabela.LoteClassificacaoProduto, "Lote_id ='" & ddlLote.SelectedValue & "' and Produto_id ='" & cmbProdutos.SelectedValue & "'")
        End If
    End Sub

    Private Sub CarregarCentrodeCusto()
        ddl.Carregar(ddlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "", True)
    End Sub

    Private Sub CarregarAnalises()
        Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
        Dim strDeposito As String() = txtCodigoDeposito.Value.Split("-")
        Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")
        Dim strMovimento As String = Convert.ToDateTime(txtMovimento.Text).ToString("yyyy-MM-dd")

        Dim strSQL As String = "SELECT PA.Analise_Id, A.Descricao, A.IndiceMinimo, A.IndiceMaximo, NULL AS Indice " &
                               "FROM ProdutosXAnalises PA " &
                               "INNER JOIN Analises A " &
                               "ON A.Analise_Id = PA.Analise_Id " &
                               "WHERE PA.Produto_Id = '" & cmbProdutos.SelectedValue & "' " &
                               "AND PA.Analise_Id NOT IN (SELECT PA2.Analise_Id " &
                                                          "FROM ProducaoXAnalises PA2 " &
                                                          "INNER JOIN Analises A " &
                                                          "ON A.Analise_Id = PA2.Analise_Id " &
                                                          "WHERE PA2.Empresa_Id = '" & strEmpresa(0) & "' " &
                                                          "AND PA2.EndEmpresa_Id = " & strEmpresa(1) & " " &
                                                          "AND PA2.Deposito_Id = '" & strDeposito(0) & "' " &
                                                          "AND PA2.EndDeposito_Id = " & strDeposito(1) & " " &
                                                          "AND PA2.Produto_Id = '" & cmbProdutos.SelectedValue & "' " &
                                                          "AND PA2.Operacao_Id = " & strOperacao(0) & " " &
                                                          "AND PA2.SubOperacao_Id = " & strOperacao(1) & " " &
                                                          "AND PA2.Movimento_Id = '" & strMovimento & "' " &
                                                          "AND PA2.FisicoFiscal_Id = " & IIf(chkFisico.Checked, "1", "2") & " " &
                                                          "AND PA2.ProdutoDerivado_Id = '" & cmbProdutosDerivados.SelectedValue & "') " &
                               "UNION " &
                               "SELECT PA.Analise_Id, A.Descricao, A.IndiceMinimo, A.IndiceMaximo, PA.Indice " &
                               "FROM ProducaoXAnalises PA " &
                               "INNER JOIN Analises A " &
                               "ON A.Analise_Id = PA.Analise_Id " &
                               "WHERE PA.Empresa_Id = '" & strEmpresa(0) & "' " &
                               "AND PA.EndEmpresa_Id = " & strEmpresa(1) & " " &
                               "AND PA.Deposito_Id = '" & strDeposito(0) & "' " &
                               "AND PA.EndDeposito_Id = " & strDeposito(1) & " " &
                               "AND PA.Produto_Id = '" & cmbProdutos.SelectedValue & "' " &
                               "AND PA.Operacao_Id = " & strOperacao(0) & " " &
                               "AND PA.SubOperacao_Id = " & strOperacao(1) & " " &
                               "AND PA.Movimento_Id = '" & strMovimento & "' " &
                               "AND PA.FisicoFiscal_Id = " & IIf(chkFisico.Checked, "1", "2") & " " &
                               "AND PA.ProdutoDerivado_Id = '" & cmbProdutosDerivados.SelectedValue & "' " &
                               "ORDER BY 1"

        Dim dsAnalises As DataSet = Banco.ConsultaDataSet(strSQL, "ProducaoXAnalises")

        If dsAnalises.Tables(0).Rows().Count > 0 Then
            idAnalises.Visible = True
        Else
            idAnalises.Visible = False
        End If

        gridAnalises.DataSource = dsAnalises
        gridAnalises.DataBind()

        Session("dsAnalises" & HID.Value) = dsAnalises
    End Sub

    Private Sub HabilitarCampos(ByVal Habilitar As Boolean)
        'cmbUnidade.Enabled = Habilitar
        'cmbEmpresa.Enabled = Habilitar
        cmdConsultaDeposito.Enabled = Habilitar
        txtMovimento.ReadOnly = Not Habilitar
        'chkFisico.Enabled = Habilitar
        'chkFiscal.Enabled = Habilitar
        cmbEtapas.Enabled = Habilitar
        cmdOperacao.Enabled = Habilitar
        cmbGruposProdutos.Enabled = Habilitar
        cmbProdutos.Enabled = Habilitar
        cmbGruposProdutosDerivados.Enabled = Habilitar
        cmbProdutosDerivados.Enabled = Habilitar
        ddlLote.Enabled = Habilitar
        ddlEmbalagem.Enabled = Habilitar
        ddlLoteClassificacao.Enabled = Habilitar
    End Sub

    Private Sub ApenasNumeros(ByVal Controle As TextBox)
        Controle.Attributes.Add("onkeypress", "return ValidarNumerico(this, event);")
    End Sub

    Private Function IncluirProducao(ByVal FisicoFiscal As String) As ArrayList
        Dim objSQLArray As New ArrayList()
        Dim strSQL As String

        Dim strEmbalagem As String() = ddlEmbalagem.SelectedValue.Split(";")
        Dim Emb, TEmb, QtdeEmb As String
        If ddlEmbalagem.Items.Count = 0 Or ddlEmbalagem.SelectedIndex = 0 Then
            Emb = "0"
            TEmb = ""
            QtdeEmb = "0.00000"
        Else
            Emb = strEmbalagem(0)
            TEmb = strEmbalagem(1)
            QtdeEmb = strEmbalagem(2)
        End If

        Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
        Dim strDeposito As String() = txtCodigoDeposito.Value.Replace(";", "-").Split("-")
        Dim strProduto As String = cmbProdutos.SelectedValue
        Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")
        Dim strMovimento As String = Convert.ToDateTime(txtMovimento.Text).ToString("yyyy-MM-dd")
        If txtMovimentoHora.Text.Replace(":", "").Replace("_", "").Length > 0 Then strMovimento &= " " & txtMovimentoHora.Text
        Dim strProdutoDerivado As String = cmbProdutosDerivados.SelectedValue
        Dim strEtapa As String = cmbEtapas.SelectedValue
        Dim strQuantidade As String = Str(CDbl(txtQuantidade.Text))
        Dim strEntradas As String = IIf(strOperacao(2) = "E", strQuantidade, "0")
        Dim strSaidas As String = IIf(strOperacao(2) = "S", strQuantidade, "0")

        Prod = New [Lib].Negocio.Produto(cmbProdutos.SelectedValue)

        Dim numLote As String = String.Empty
        'Dim dataFabricacao As String = String.Empty
        'Dim dataValidade As String = String.Empty

        If Prod.ControlarLote Then numLote = ddlLote.SelectedValue

        If Prod.ControlarNumeroDoLote Then
            numLote = LTrim(txtNumeroDoLote.Text)
            numLote = RTrim(numLote)
            'dataFabricacao = Format(txtDataFabricacao.Text, "yyyy-MM-dd")
            'dataValidade = Format(txtDataValidade.Text, "yyyy-MM-dd")
        End If

        strSQL = "INSERT INTO Producao " &
                 "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " &
                 "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " &
                 " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, Fabricado, Validade, CentroDeCusto) " &
                 "VALUES ('" & strEmpresa(0) & "', " & strEmpresa(1) & ", '" & strDeposito(0) & "', " & strDeposito(1) & ", " &
                 "'" & strProduto & "', " & strOperacao(0) & ", " & strOperacao(1) & ", '" & strMovimento & "', " &
                 FisicoFiscal & ", '" & strProdutoDerivado & "','" & numLote & "','" & ddlLoteClassificacao.SelectedValue & "'," &
                 Emb & ",'" & TEmb & "'," & QtdeEmb & "," &
                 strEtapa & ", 'NENHUM', " & strEntradas & ", " & strSaidas & ",'" & txtObservacao.Text & "', '" & UsuarioServidor.NomeUsuario & "', GETDATE(),"

        If numLote.Length > 0 Then
            strSQL &= "'" & CDate(txtDataFabricacao.Text).ToString("yyyy-MM-dd") & "','" & CDate(txtDataValidade.Text).ToString("yyyy-MM-dd") & "',"
        Else
            strSQL &= "NULL,NULL,"
        End If

        If ddlCentroDeCusto.SelectedValue Is Nothing Then

            strSQL &= "NULL)"

        Else
            strSQL &= "'" & ddlCentroDeCusto.SelectedValue & "')"

        End If

        objSQLArray.Add(strSQL)

        For Each row As GridViewRow In gridAnalises.Rows
            Dim lblAnalise As Label = row.FindControl("lblCodigo")
            Dim txtIndice As TextBox = row.FindControl("txtAnalise")

            If txtIndice.Text.Replace("_", "").Replace(",", "").Length > 0 Then
                strSQL = "INSERT INTO ProducaoXAnalises " &
                         "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, " &
                         "Operacao_Id, SubOperacao_Id, Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, " &
                         "Analise_Id, Etapa, Quantidade, Indice) " &
                         "VALUES ('" & strEmpresa(0) & "', " & strEmpresa(1) & ", '" & strDeposito(0) & "', " &
                         strDeposito(1) & ", '" & strProduto & "', " & strOperacao(0) & ", " & strOperacao(1) & ", " &
                         "'" & strMovimento & "', " & FisicoFiscal & ", '" & strProdutoDerivado & "','" & ddlLote.SelectedValue & "','" & ddlLoteClassificacao.SelectedValue & "'," & Emb & ",'" & TEmb & "'," & QtdeEmb & "," &
                         lblAnalise.Text & ", " & strEtapa & ", " & strQuantidade & ", " &
                         txtIndice.Text.Replace(",", ".") & ") "

                objSQLArray.Add(strSQL)
            End If
        Next

        Return objSQLArray
    End Function
#End Region

#Region "Tab Consulta"

    Protected Sub grdConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dsProducao As DataSet = CType(Session("dsProducao" & HID.Value), DataSet)
        Dim drProducao As DataRow = dsProducao.Tables(0).Rows(grdConsulta.SelectedIndex)

        cmbUnidade.SelectedIndex = -1
        Dim itemUnidade As ListItem = cmbUnidade.Items.FindByValue(drProducao("GrupoEmpresa").ToString())
        itemUnidade.Selected = True
        CarregarEmpresas()

        cmbEmpresa.SelectedIndex = -1

        For Each itemEmpresa As ListItem In cmbEmpresa.Items
            Dim strEmpresa As String() = itemEmpresa.Value.Split("-")

            If strEmpresa.GetUpperBound(0) > 0 Then
                If strEmpresa(0) = drProducao("CNPJEmpresa").ToString() And strEmpresa(1) = drProducao("EndEmpresa_Id").ToString() Then
                    itemEmpresa.Selected = True
                    Exit For
                End If
            End If
        Next

        txtDeposito.Text = drProducao("NomeDeposito").ToString() & " - " & drProducao("CidadeDeposito").ToString() & " - " & drProducao("EstadoDeposito").ToString() & " - " & Funcoes.FormatarCpfCnpj(drProducao("CNPJDeposito").ToString())
        txtCodigoDeposito.Value = drProducao("CNPJDeposito").ToString() & "-" & drProducao("EndDeposito_Id").ToString()

        txtMovimento.Text = Convert.ToDateTime(drProducao("Movimento")).ToString("dd/MM/yyyy")
        txtMovimentoHora.Text = Convert.ToDateTime(drProducao("Movimento")).ToString("HH:mm")

        If drProducao("Tipo").ToString() = "Fisico" Then
            chkFisico.Checked = True
            chkFiscal.Checked = False
        Else
            chkFiscal.Checked = True
            chkFisico.Checked = False
        End If

        cmbEtapas.SelectedIndex = -1
        Dim itemEtapa As ListItem = cmbEtapas.Items.FindByValue(drProducao("Etapa").ToString())
        itemEtapa.Selected = True

        txtOperacao.Text = drProducao("Operacao_Id").ToString() & "-" & drProducao("SubOperacao_Id").ToString() & " - " & drProducao("NomeOperacao").ToString()
        txtCodigoOperacao.Value = drProducao("Operacao_Id").ToString() & ";" & drProducao("SubOperacao_Id").ToString() & ";" &
                                  drProducao("EntradaSaida").ToString() & ";" & drProducao("Flags").ToString()

        cmbGruposProdutos.SelectedIndex = -1
        Dim itemGrupoProduto As ListItem = cmbGruposProdutos.Items.FindByValue(drProducao("Grupo").ToString())
        itemGrupoProduto.Selected = True
        CarregarProdutos("Normal")

        cmbProdutos.SelectedIndex = -1
        Dim itemProduto As ListItem = cmbProdutos.Items.FindByValue(drProducao("Produto_Id").ToString())
        itemProduto.Selected = True

        cmbGruposProdutosDerivados.SelectedIndex = -1
        If drProducao("GrupoDerivado").ToString.Length > 0 Then
            Dim itemGrupoProdutoDerivado As ListItem = cmbGruposProdutosDerivados.Items.FindByValue(drProducao("GrupoDerivado").ToString())
            itemGrupoProdutoDerivado.Selected = True
            CarregarProdutos("Derivado")
        End If

        cmbProdutosDerivados.SelectedIndex = -1
        If drProducao("ProdutoDerivado_Id").ToString.Length > 0 Then
            Dim itemProdutoDerivado As ListItem = cmbProdutosDerivados.Items.FindByValue(drProducao("ProdutoDerivado_Id").ToString())
            itemProdutoDerivado.Selected = True
        End If

        ddlLote.Items.Clear()
        ddlLote.Items.Add(New ListItem(drProducao("Lote_Id"), drProducao("Lote_Id")))
        ddlLote.SelectedIndex = 0

        ddlLoteClassificacao.Items.Clear()
        ddlLoteClassificacao.Items.Add(New ListItem(drProducao("Classificacao_Id"), drProducao("Classificacao_Id")))
        ddlLoteClassificacao.SelectedIndex = 0

        ddlEmbalagem.Items.Clear()
        ddlEmbalagem.Items.Add(New ListItem(drProducao("DDLDescricaoEmbalagem"), drProducao("DDLCodigoEmbalagem")))
        ddlEmbalagem.SelectedIndex = 0

        txtQuantidade.Text = CDbl(drProducao("Quantidade")).ToString("N4")
        txtObservacao.Text = drProducao("Observacao").ToString

        Dim prd As New Produto(drProducao("Produto_Id"))
        If prd.ControlarNumeroDoLote Then
            idControlarNumeroDoLote.Visible = True
            imgSelecionaLote.Visible = True
            txtNumeroDoLote.Text = drProducao("Lote_Id")

            Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")

            If strOperacao(2) = "E" Then
                imgSelecionaLote.Visible = False
            End If

            If drProducao("OrdemDeProducao") = 0 Then
                txtDataFabricacao.Enabled = True
                txtDataValidade.Enabled = True
                txtNumeroDoLote.Enabled = True
            End If

            txtDataFabricacao.Text = CDate(drProducao("Fabricado")).ToString("dd/MM/yyyy")
            txtDataValidade.Text = CDate(drProducao("Validade")).ToString("dd/MM/yyyy")
        Else
            idControlarNumeroDoLote.Visible = False
            imgSelecionaLote.Visible = False
            txtNumeroDoLote.Text = String.Empty
        End If

        CarregarAnalises()
        HabilitarCampos(False)

        tabcProducao.ActiveTabIndex = 0

        If drProducao("OrdemDeProducao") > 0 Then
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            txtNumeroDoLote.Enabled = False

            MsgBox(Me.Page, "Registro gerado pela Ordem de Produção " & drProducao("OrdemDeProducao").ToString & ", apenas Consulta.", eTitulo.Info)
        Else
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        End If
    End Sub

#End Region

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ControleDeProducao", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim objSqlArray As New ArrayList
                    If chkFisico.Checked Then objSqlArray.AddRange(IncluirProducao("1"))
                    If chkFiscal.Checked Then objSqlArray.AddRange(IncluirProducao("2"))
                    If Banco.GravaBanco(objSqlArray) Then
                        MsgBox(Me.Page, "sucesso na inclusão.", eTitulo.Sucess)
                        If Not chkManterDados.Checked Then LimparCampos()
                    Else
                        MsgBox(Me.Page, "Erro ao Incluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ControleDeProducao", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim objSQL As New ArrayList
                    Dim Emb, TEmb, QtdeEmb As String
                    If ddlEmbalagem.SelectedValue = "" Then
                        Emb = "0"
                        TEmb = ""
                        QtdeEmb = "0.00000"
                    Else
                        Dim strEmbalagem As String() = ddlEmbalagem.SelectedValue.Split(";")
                        Emb = strEmbalagem(0)
                        TEmb = strEmbalagem(1)
                        QtdeEmb = strEmbalagem(2)
                    End If

                    Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
                    Dim strDeposito As String() = txtCodigoDeposito.Value.Split("-")
                    Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")
                    Dim strMovimento As String = Convert.ToDateTime(txtMovimento.Text).ToString("yyyy-MM-dd")
                    Dim strQuantidade As String = txtQuantidade.Text.Replace(".", "")

                    Prod = New [Lib].Negocio.Produto(cmbProdutos.SelectedValue)

                    Dim numLote As String = String.Empty

                    If Prod.ControlarLote Then numLote = ddlLote.SelectedValue

                    If Prod.ControlarNumeroDoLote Then
                        numLote = LTrim(txtNumeroDoLote.Text)
                        numLote = RTrim(numLote)
                    End If

                    Dim strSQL As String = "UPDATE Producao " &
                                           "   SET " & IIf(strOperacao(2) = "E", "Entradas", "Saidas") & " = " & Str(strQuantidade) & ", " & vbCrLf

                    If Prod.ControlarNumeroDoLote Then
                        strSQL &= "       Fabricado              ='" & CDate(txtDataFabricacao.Text).ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                                  "       Validade               ='" & CDate(txtDataValidade.Text).ToString("yyyy-MM-dd") & "', " & vbCrLf
                    End If

                    strSQL &= "       Observacao             ='" & txtObservacao.Text & "' " & vbCrLf & _
                                           " WHERE Empresa_Id             ='" & strEmpresa(0) & "' " & vbCrLf & _
                                           "   AND EndEmpresa_Id          = " & strEmpresa(1) & " " & vbCrLf & _
                                           "   AND Deposito_Id            ='" & strDeposito(0) & "' " & vbCrLf & _
                                           "   AND EndDeposito_Id         = " & strDeposito(1) & " " & vbCrLf & _
                                           "   AND Produto_Id             ='" & cmbProdutos.SelectedValue & "' " & vbCrLf & _
                                           "   AND Operacao_Id            = " & strOperacao(0) & " " & vbCrLf & _
                                           "   AND SubOperacao_Id         = " & strOperacao(1) & " " & vbCrLf & _
                                           "   AND Movimento_Id           ='" & strMovimento & "' " & vbCrLf & _
                                           "   AND FisicoFiscal_Id        = " & IIf(chkFisico.Checked, "1", "2") & " " & vbCrLf & _
                                           "   AND ProdutoDerivado_Id     ='" & cmbProdutosDerivados.SelectedValue & "'" & vbCrLf & _
                                           "   AND Lote_Id                ='" & numLote & "'" & vbCrLf & _
                                           "   AND Classificacao_Id       ='" & ddlLoteClassificacao.SelectedValue & "'" & vbCrLf & _
                                           "   AND Embalagem_Id           = " & Emb & vbCrLf & _
                                           "   AND TipoDeEmbalagem_Id     ='" & TEmb & "'" & vbCrLf & _
                                           "   AND CapacidadeEmbalagem_Id = " & QtdeEmb & vbCrLf
                    objSQL.Add(strSQL)

                    strSQL = "DELETE ProducaoXAnalises " & vbCrLf & _
                             " WHERE Empresa_Id         ='" & strEmpresa(0) & "'" & vbCrLf & _
                             "   AND EndEmpresa_Id      = " & strEmpresa(1) & vbCrLf & _
                             "   AND Deposito_Id        ='" & strDeposito(0) & "'" & vbCrLf & _
                             "   AND EndDeposito_Id     = " & strDeposito(1) & vbCrLf & _
                             "   AND Produto_Id         ='" & cmbProdutos.SelectedValue & "' " & vbCrLf & _
                             "   AND Operacao_Id        = " & strOperacao(0) & vbCrLf & _
                             "   AND SubOperacao_Id     = " & strOperacao(1) & vbCrLf & _
                             "   AND Movimento_Id       ='" & strMovimento & "'" & vbCrLf & _
                             "   AND FisicoFiscal_Id    = " & IIf(chkFisico.Checked, "1", "2") & vbCrLf & _
                             "   AND ProdutoDerivado_Id ='" & cmbProdutosDerivados.SelectedValue & "'" & vbCrLf & _
                             "   AND Lote_Id            ='" & ddlLote.SelectedValue & "'" & vbCrLf & _
                             "   AND Classificacao_Id   ='" & ddlLoteClassificacao.SelectedValue & "'" & vbCrLf & _
                             "   AND Embalagem_Id           = " & Emb & vbCrLf & _
                             "   AND TipoDeEmbalagem_Id     ='" & TEmb & "'" & vbCrLf & _
                             "   AND CapacidadeEmbalagem_Id = " & QtdeEmb & vbCrLf

                    objSQL.Add(strSQL)

                    For Each row As GridViewRow In gridAnalises.Rows
                        Dim lblAnalise As Label = row.FindControl("lblCodigo")
                        Dim txtIndice As TextBox = row.FindControl("txtAnalise")

                        If txtIndice.Text.Replace("_", "").Replace(",", "").Length > 0 Then
                            strSQL = "INSERT INTO ProducaoXAnalises " & vbCrLf & _
                                     "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, " & vbCrLf & _
                                     "Operacao_Id, SubOperacao_Id, Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id," & vbCrLf & _
                                     "Analise_Id, Etapa, Quantidade, Indice) " & vbCrLf & _
                                     "VALUES ('" & strEmpresa(0) & "', " & strEmpresa(1) & ", '" & strDeposito(0) & "', " & vbCrLf & _
                                     strDeposito(1) & ", '" & cmbProdutos.SelectedValue & "', " & strOperacao(0) & ", " & vbCrLf & _
                                     strOperacao(1) & ", '" & strMovimento & "', " & IIf(chkFisico.Checked, "1", "2") & ", " & vbCrLf & _
                                     "'" & cmbProdutosDerivados.SelectedValue & "','" & ddlLote.SelectedValue & "','" & ddlLoteClassificacao.SelectedValue & "'," & Emb & ",'" & TEmb & "'," & QtdeEmb & ", " & lblAnalise.Text & ", " & vbCrLf & _
                                     cmbEtapas.SelectedValue & ", " & strQuantidade & ", " & txtIndice.Text.Replace(",", ".") & ") "

                            objSQL.Add(strSQL)
                        End If
                    Next

                    If Banco.GravaBanco(objSQL) Then
                        MsgBox(Me.Page, "sucesso na alteração.", eTitulo.Sucess)
                        If Not chkManterDados.Checked Then LimparCampos()
                    Else
                        MsgBox(Me.Page, "Erro ao Alterar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ControleDeProducao", "EXCLUIR") Then
                If ValidarCampos() Then
                    Dim objSQL As New ArrayList
                    Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
                    Dim strDeposito As String() = txtCodigoDeposito.Value.Split("-")
                    Dim strOperacao As String() = txtCodigoOperacao.Value.Split(";")
                    Dim strMovimento As String = Convert.ToDateTime(txtMovimento.Text).ToString("yyyy-MM-dd")

                    Dim strEmbalagem As String() = ddlEmbalagem.SelectedValue.Split(";")
                    Dim Emb, TEmb, QtdeEmb As String
                    If ddlEmbalagem.SelectedValue.Length = 0 Then
                        Emb = "0"
                        TEmb = ""
                        QtdeEmb = "0.00000"
                    Else
                        Emb = strEmbalagem(0)
                        TEmb = strEmbalagem(1)
                        QtdeEmb = strEmbalagem(2)
                    End If

                    Dim strSQL As String = "DELETE ProducaoXAnalises " &
                                           " WHERE Empresa_Id         ='" & strEmpresa(0) & "' " &
                                           "   AND EndEmpresa_Id      = " & strEmpresa(1) & " " &
                                           "   AND Deposito_Id        ='" & strDeposito(0) & "' " &
                                           "   AND EndDeposito_Id     = " & strDeposito(1) & " " &
                                           "   AND Produto_Id         ='" & cmbProdutos.SelectedValue & "' " &
                                           "   AND Operacao_Id        = " & strOperacao(0) & " " &
                                           "   AND SubOperacao_Id     = " & strOperacao(1) & " " &
                                           "   AND Movimento_Id       ='" & strMovimento & "' " &
                                           "   AND ProdutoDerivado_Id ='" & cmbProdutosDerivados.SelectedValue & "'" &
                                           "   AND Lote_Id                ='" & ddlLote.SelectedValue & "'" &
                                           "   AND Classificacao_Id       ='" & ddlLoteClassificacao.SelectedValue & "'" &
                                           "   AND Embalagem_Id           = " & Emb &
                                           "   AND TipoDeEmbalagem_id     ='" & TEmb & "'" &
                                           "   AND CapacidadeEmbalagem_Id = " & QtdeEmb
                    If Not chkFisico.Checked Or Not chkFiscal.Checked Then
                        strSQL &= "AND FisicoFiscal_Id = " & IIf(chkFisico.Checked, "1", "2") & " "
                    End If

                    objSQL.Add(strSQL)

                    strSQL = "DELETE Producao " &
                             " WHERE Empresa_Id             ='" & strEmpresa(0) & "' " &
                             "   AND EndEmpresa_Id          = " & strEmpresa(1) & " " &
                             "   AND Deposito_Id            ='" & strDeposito(0) & "' " &
                             "   AND EndDeposito_Id         = " & strDeposito(1) & " " &
                             "   AND Produto_Id             ='" & cmbProdutos.SelectedValue & "' " &
                             "   AND Operacao_Id            = " & strOperacao(0) & " " &
                             "   AND SubOperacao_Id         = " & strOperacao(1) & " " &
                             "   AND Movimento_Id           ='" & strMovimento & "' " &
                             "   AND ProdutoDerivado_Id     ='" & cmbProdutosDerivados.SelectedValue & "'" &
                             "   AND Lote_Id                ='" & ddlLote.SelectedValue & "'" &
                             "   AND Classificacao_Id       ='" & ddlLoteClassificacao.SelectedValue & "'" &
                             "   AND Embalagem_Id           = " & Emb &
                             "   AND TipoDeEmbalagem_id     ='" & TEmb & "'" &
                             "   AND CapacidadeEmbalagem_Id = " & QtdeEmb
                    If Not chkFisico.Checked Or Not chkFiscal.Checked Then
                        strSQL &= "AND FisicoFiscal_Id = " & IIf(chkFisico.Checked, "1", "2") & " "
                    End If

                    objSQL.Add(strSQL)

                    If Banco.GravaBanco(objSQL) Then
                        MsgBox(Me.Page, "sucesso na exclusão.", eTitulo.Sucess)
                        LimparCampos()
                    Else
                        MsgBox(Me.Page, "Erro ao Excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProduto_Click(sender As Object, e As EventArgs) Handles lnkBuscaProduto.Click
        Try
            If txtCodigoOperacao.Value.Length = 0 Then
                MsgBox(Me.Page, "Operação deve ser selecionada antes.")
                Exit Sub
            End If

            Session("Where" & HID.Value) = " Situacao = 1 "
            ucConsultaProduto.Limpar()
            ucConsultaProduto.BuscarProduto()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoETQ" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgSelecionaLote_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        If String.IsNullOrWhiteSpace(txtQuantidade.Text) Then
            MsgBox(Me.Page, "Quantidade não foi informada", eTitulo.Info)
            Exit Sub
        ElseIf CDec(txtQuantidade.Text) = 0 Then
            MsgBox(Me.Page, "Quantidade não foi informada", eTitulo.Info)
            Exit Sub
        End If

        Dim ucConsultaLote = CType(Me.Page.FindControlRecursive("ucConsultaLote"), ucConsultaLote)
        If ucConsultaLote IsNot Nothing Then
            ucConsultaLote.Limpar()
            ucConsultaLote.SetarHID(HID.Value)


            Dim nf As NotaFiscal = New NotaFiscal()
            Dim item As NotaFiscalXItem = New NotaFiscalXItem()
            item.NotaFiscal = nf
            item.NotaFiscal.CodigoEmpresa = cmbEmpresa.SelectedValue.Split("-")(0)
            item.NotaFiscal.EnderecoEmpresa = cmbEmpresa.SelectedValue.Split("-")(1)
            'item.NotaFiscal.CodigoRomaneio = 0
            item.CodigoProduto = cmbProdutos.SelectedValue
            item.QuantidadeFiscal = CDec(txtQuantidade.Text)


            ucConsultaLote.CarregarLote(item.NotaFiscal.CodigoEmpresa, item.NotaFiscal.EnderecoEmpresa, item)
        End If

        Popup.ConsultaDeLote(Me.Page, "objConsultaDeLote" & HID.Value)
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        LimparCampos()
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        If cmbEmpresa.SelectedIndex > 0 Or txtCodigoDeposito.Value.Length > 0 Or txtMovimento.Text.Length > 0 Then
            tabcProducao.ActiveTabIndex = 1
            ConsultarProducao()
        Else
            MsgBox(Me.Page, "Informe um dos seguintes campos para consulta:\n-Empresa\n-Depósito\n-Data de movimento")
            LimparCampos()
        End If
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ControleDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(cmbUnidade.SelectedValue) Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        End If
        If String.IsNullOrWhiteSpace(cmbEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ConsistenciaDeProducao", "RELATORIO") Then
                Dim sql As String = ""

                If Validar() Then

                    sql = "SELECT Clientes.Reduzido, Clientes.Nome AS NomeDoDeposito, " & vbCrLf &
                          "       Clientes.Cidade, Clientes.Estado, " & vbCrLf &
                          "       Producao.Produto_Id AS Produto, Produtos.Nome AS NomeDoProduto, " & vbCrLf &
                          "       Producao.Operacao_Id AS Operacao, Producao.SubOperacao_Id AS SubOperacao, SubOperacoes.Descricao AS NomeDaOperacao, " & vbCrLf &
                          "       Producao.Movimento_Id AS Movimento, Producao.Entradas, Producao.Saidas, case when Producao.FisicoFiscal_Id = 1 THEN 'Fisico' ELSE 'Fiscal' end AS FisicoFiscal," & vbCrLf &
                          "       ISNULL(ProducaoXAnalises.Analise_Id, 0) AS Analise, ISNULL(Analises.Descricao, N'') AS NomeDaAnalise, ISNULL(ProducaoXAnalises.Indice, 0) AS Indice, CentroDeCusto" & vbCrLf &
                          "       ,Producao.Lote_Id + case when len(Producao.Lote_Id)>0 then '-' else '' end +Producao.Classificacao_Id +'('+ Embalagens.EmbalagemIndea +'-'+ Producao.TipoDeEmbalagem_Id+' / '+ Convert(nvarchar(10),isnull(Producao.CapacidadeEmbalagem_Id,''))+')' as LoteEmbalagem, Producao.ProdutoDerivado_Id as ProdutoDerivado" & vbCrLf &
                          "  FROM Clientes" & vbCrLf &
                          " INNER JOIN Producao" & vbCrLf

                    If Not CkCDeposito.Checked Then
                        sql &= "    ON Clientes.Cliente_Id  = Producao.Deposito_Id " & vbCrLf &
                               "   AND Clientes.Endereco_Id = Producao.EndDeposito_Id " & vbCrLf
                    Else
                        sql &= "    ON Clientes.Cliente_Id  = Producao.Empresa_Id " & vbCrLf &
                               "   AND Clientes.Endereco_Id = Producao.EndEmpresa_Id " & vbCrLf
                    End If

                    sql &= " INNER JOIN Produtos " & vbCrLf &
                          "    ON Producao.Produto_Id = Produtos.Produto_Id " & vbCrLf &
                          " INNER JOIN SubOperacoes " & vbCrLf &
                          "    ON Producao.Operacao_Id    = SubOperacoes.Operacao_Id " & vbCrLf &
                          "   AND Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                          "  LEFT OUTER JOIN ProducaoXAnalises " & vbCrLf &
                          "    ON Producao.Empresa_Id             = ProducaoXAnalises.Empresa_Id " & vbCrLf &
                          "   AND Producao.EndEmpresa_Id          = ProducaoXAnalises.EndEmpresa_Id " & vbCrLf &
                          "   AND Producao.Deposito_Id            = ProducaoXAnalises.Deposito_Id " & vbCrLf &
                          "   AND Producao.EndDeposito_Id         = ProducaoXAnalises.EndDeposito_Id " & vbCrLf &
                          "   AND Producao.Produto_Id             = ProducaoXAnalises.Produto_Id " & vbCrLf &
                          "   AND Producao.Operacao_Id            = ProducaoXAnalises.Operacao_Id " & vbCrLf &
                          "   AND Producao.SubOperacao_Id         = ProducaoXAnalises.SubOperacao_Id " & vbCrLf &
                          "   AND Producao.Movimento_Id           = ProducaoXAnalises.Movimento_Id " & vbCrLf &
                          "   AND Producao.FisicoFiscal_Id        = ProducaoXAnalises.FisicoFiscal_Id " & vbCrLf &
                          "   AND Producao.ProdutoDerivado_Id     = ProducaoXAnalises.ProdutoDerivado_Id" & vbCrLf &
                          "   And Producao.Lote_Id                = ProducaoXAnalises.Lote_Id" & vbCrLf &
                          "   And Producao.Classificacao_Id       = ProducaoXAnalises.Classificacao_Id" & vbCrLf &
                          "   And Producao.Embalagem_Id           = ProducaoXAnalises.Embalagem_Id" & vbCrLf &
                          "   And Producao.TipoDeEmbalagem_Id     = ProducaoXAnalises.TipoDeEmbalagem_Id" & vbCrLf &
                          "   And Producao.CapacidadeEmbalagem_Id = ProducaoXAnalises.CapacidadeEmbalagem_Id" & vbCrLf &
                          "  LEFT OUTER JOIN Analises " & vbCrLf &
                          "    ON Analises.Analise_Id = ProducaoXAnalises.Analise_Id " & vbCrLf &
                          "  Left Join Embalagens" & vbCrLf &
                          "    on Embalagens.Embalagem_id = Producao.Embalagem_Id" & vbCrLf

                    sql &= " WHERE Producao.Empresa_Id = '" & cmbEmpresa.SelectedValue.Split("-")(0) & "' And Producao.EndEmpresa_Id = " & cmbEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                           "   AND Producao.Movimento_Id BETWEEN '" & txtMovimento.Text.ToSqlDate() & "' " & vbCrLf &
                           "   AND '" & Convert.ToDateTime(txtMovimento.Text).AddSeconds(86399).ToString("yyyy-MM-dd HH:mm:ss") & "' " & vbCrLf

                    If Not String.IsNullOrWhiteSpace(cmbProdutos.SelectedValue) Then
                        sql &= " And (Producao.Produto_Id = '" & cmbProdutos.SelectedValue & "')"
                    End If

                    If chkFisico.Checked = True And chkFiscal.Checked = False Then
                        sql &= " And (Producao.FisicoFiscal_Id = 1)"
                    End If

                    If chkFiscal.Checked = True Then
                        If chkFisico.Checked = True Then
                            sql &= " And (Producao.FisicoFiscal_Id = 1 Or Producao.FisicoFiscal_Id = 2)"
                        Else
                            sql &= " And (Producao.FisicoFiscal_Id = 2)"
                        End If
                    End If

                    Dim DS As DataSet = Banco.ConsultaDataSet(sql, "ConsistenciaDeProducao")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Titulo", "Relatório De Consistência De Produção.")
                    parameters.Add("ConsultaParametros", "")

                    Funcoes.BindReport(Me.Page, DS, "Cr_ConsistenciaDeProducao", eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class