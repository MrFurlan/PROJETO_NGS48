Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucNFEncargo
    Inherits BaseUserControl

#Region "Variáveis"

    Private Property objNotaFiscal As [Lib].Negocio.NotaFiscal
        Get
            Return CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
        End Get
        Set(value As [Lib].Negocio.NotaFiscal)
            Session("objNotaFiscal" & HID.Value) = CType(value, [Lib].Negocio.NotaFiscal)
        End Set
    End Property

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

    'Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private dblEncargos As Double
    Private dblLiquido As Double
    Private dblTotal As Double

#End Region

#Region "Eventos"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            CarregarProdutos()
        End If
    End Sub

    Protected Sub lnkRecarregar_Click(sender As Object, e As EventArgs) Handles lnkRecarregar.Click
        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).IUD = "U"

        Dim OxE As New OperacaoXEstado()
        OxE.Codigo = ddlVersao.SelectedValue
        OxE = New OperacaoXEstado(OxE)

        If Not OxE.Ativo Then
            MsgBox(Me.Page, "Atenção, ID selecionada está CANCELADA.", eTitulo.Info)
            Exit Sub
        End If

        If OxE.InicioVigencia > objNotaFiscal.Movimento Then
            MsgBox(Me.Page, "Atenção, ID selecionada não pode ter a Vigência maior que a Data do Movimento da Nota.", eTitulo.Info)
            Exit Sub
        End If

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

        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoOperacaoEstado = ddlVersao.SelectedValue

        objNotaFiscal.CarregandoNota = True
        Dim nXe = New ListNotaFiscalXItemXEncargo(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})

        If nXe.Count > 0 Then
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos = Nothing
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos = nXe

            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoOperacao = nXe.EncProduto.CodigoOperacao
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoSubOperacao = nXe.EncProduto.CodigoSubOperacao

            If hdCentroDeCusto.Value > 0 Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos.EncProduto.CentroDeCusto = hdCentroDeCusto.Value

            DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
            DgEncargos.DataBind()
            HabilitarCampos()

            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                objNotaFiscal.TipoDeDocumentoFrete = Nothing

                objNotaFiscal.CodigoCliente = fCliente
                objNotaFiscal.EnderecoCliente = fEndCliente
            End If
        Else
            MsgBox(Me.Page, "Encargos não foram encotrados/recarregados")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        If TypeOf Me.Page Is NotasFiscaisGerais Then
            CType(Me.Page, NotasFiscaisGerais).AtualizarItensNoGrid()
            CType(Me.Page, NotasFiscaisGerais).AtualizarValorParcelamento()
        ElseIf TypeOf Me.Page Is NotaFiscalXItens Then
            'CType(Me.Page, NotaFiscalXItens).AtualizaFormularioComAClasse()
            CType(Me.Page, NotaFiscalXItens).AtualizaComEncargos()
        End If
        Popup.CloseDialog(Me.Page, "divNFEncargo")
    End Sub

    Protected Sub ddlProdutoSel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutoSel.SelectedIndexChanged
        hdfIndex.Value = ddlProdutoSel.SelectedIndex
        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
        DgEncargos.DataBind()

        Dim Parametros As New OperacaoXEstado
        Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
        Parametros.CodigoGrupoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Produto.CodigoGrupo
        Parametros.CodigoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoProduto
        Parametros.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoOperacao
        Parametros.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoSubOperacao
        Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado
        Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado

        If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
            If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                Parametros.EstadoDestino = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado
            End If
        End If

        Dim ObjListVersoes As New ListOperacaoXEstado(Parametros)

        ddlVersao.Items.Clear()

        ddlVersao.DataTextField = "DESCRICAO"
        ddlVersao.DataValueField = "CODIGO"

        ddlVersao.DataSource = (From x In ObjListVersoes
                                Order By x.InicioVigencia, x.Codigo
                                Select x.Codigo, Descricao = x.InicioVigencia.ToString("dd-MM-yyyy") & "   ID-" & x.Codigo & "   " & x.UsuarioInclusao & "   " & x.DataHoraInclusao.ToString("dd-MM-yyyy") & IIf(x.Ativo, "   Ativo", "   Cancelada"))
        ddlVersao.DataBind()

        ddlVersao.SelectedValue = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).CodigoOperacaoEstado

        HabilitarCampos()
    End Sub

    Protected Sub DgEncargos_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles DgEncargos.RowCreated
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim btnEncargoItem As Button = CType(e.Row.FindControl("btnEncargoItem"), Button)
            btnEncargoItem.CommandArgument = e.Row.RowIndex.ToString()
        End If
    End Sub

    Protected Sub DgEncargos_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles DgEncargos.RowCommand
        If e.CommandName = "OK" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            AtualizarEncargos(index)
        End If
    End Sub
#End Region



#Region "Métodos"
    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        hdfIndex.Value = ""
        hdCentroDeCusto.Value = 0
        If ddlProdutoSel.Items.Count > 0 Then
            ddlProdutoSel.SelectedIndex = 0
            hdfIndex.Value = ddlProdutoSel.SelectedIndex
        End If
        lnkRecarregar.Parent.Visible = objNotaFiscal.IUD = "U"

        txtBeneficioICMS.Text = String.Empty
        linhaBeneficio.Visible = False

        'If Not String.IsNullOrWhiteSpace(hdfIndex.Value) Then
        '    lblConfigOperacao.Text = objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = ddlProdutoSel.SelectedValue And s.IUD <> "D")(hdfIndex.Value).CodigoOperacaoEstado
        'Else
        '    lblConfigOperacao.Text = String.Empty
        'End If

        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = ddlProdutoSel.SelectedValue And s.IUD <> "D").SelectMany(Function(s) s.Encargos)
        DgEncargos.DataBind()

        HabilitarCampos()
    End Sub

    Public Sub Inicializar(ByVal posicao As Integer)
        CarregarProdutos()
        hdfIndex.Value = posicao

        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).Encargos.Count > 0 Then
            hdCentroDeCusto.Value = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).Encargos.EncProduto.CentroDeCusto
        End If

        ddlProdutoSel.SelectedValue = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoProduto
        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
        DgEncargos.DataBind()

        idEncargosXml.Visible = False

        If Not SessaoDsXML Is Nothing Then
            If objNotaFiscal.IUD = "I" Then
                idEncargosXml.Visible = True
                CarregaEncargosXML(posicao)
            End If
        End If

        If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).OperacaoEstado Is Nothing AndAlso objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).OperacaoEstado.CodigoBeneficio.Length > 0 Then
            linhaBeneficio.Visible = True
            txtBeneficioICMS.Text = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).OperacaoEstado.BeneficioICMS.Codigo & " - " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).OperacaoEstado.BeneficioICMS.Descricao
        End If

        Dim Parametros As New OperacaoXEstado

        If objNotaFiscal.IUD = "U" Then
            Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
            Parametros.Codigo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoOperacaoEstado
        ElseIf Not objNotaFiscal.NossaEmissao Then
            Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
            Parametros.CodigoGrupoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).Produto.CodigoGrupo
            Parametros.CodigoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoProduto
            Parametros.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoOperacao
            Parametros.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoSubOperacao

            Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado

            Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado
            If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao

            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                    objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                    objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                    Parametros.EstadoDestino = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado
                    Parametros.RegiaoDestino = objNotaFiscal.NotasTrocaOrigem(0).Cliente.Estado.Regiao
                    'If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
                End If
            End If
        Else
            Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
            Parametros.Codigo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoOperacaoEstado
        End If

        Dim ObjListVersoes As New ListOperacaoXEstado(Parametros)

        ddlVersao.Items.Clear()

        ddlVersao.DataTextField = "DESCRICAO"
        ddlVersao.DataValueField = "CODIGO"

        ddlVersao.DataSource = (From x In ObjListVersoes
                                Order By x.InicioVigencia, x.Codigo
                                Select x.Codigo, Descricao = x.InicioVigencia.ToString("dd-MM-yyyy") & "   ID-" & x.Codigo & "   " & x.UsuarioInclusao & "   " & x.DataHoraInclusao.ToString("dd-MM-yyyy") & IIf(x.Ativo, "   Ativo", "   Cancelada"))
        ddlVersao.DataBind()

        'ddlVersao.SelectedValue = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoOperacaoEstado

        'ddlVersao.SelectedIndex = 0

        With ddlVersao
            ddlVersao.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoOperacaoEstado))
        End With

        'If objNotaFiscal.NFG AndAlso objNotaFiscal.IUD = "U" Then ddlVersao.Enabled = True
        If Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.IUD = "U" Then ddlVersao.Enabled = True

        HabilitarCampos()
    End Sub

    Private Sub CarregarProdutos()
        ddlProdutoSel.Items.Clear()
        If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.Itens IsNot Nothing Then
            For Each item As NotaFiscalXItem In objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D" And Not s.Produto Is Nothing)
                Dim row As ListItem = New ListItem(String.Format("{0} - {1}", item.CodigoProduto, item.Produto.Nome), item.CodigoProduto)
                ddlProdutoSel.Items.Add(row)
            Next
            ddlProdutoSel.Enabled = objNotaFiscal.NFG
        End If
    End Sub

    Private Sub CarregaEncargosXML(posicaoGrid As Integer)

        Dim dsXml As DataSet = SessaoDsXML
        Dim itens As ListNotaFiscalXItemXEncargo = New ListNotaFiscalXItemXEncargo()
        Dim indiceImposto As Integer = -1

        Dim dtXmlEncargos As New DataTable("EncargosXML")

        dtXmlEncargos.Columns.Add("Codigo", GetType(String))
        dtXmlEncargos.Columns.Add("SituacaoTributaria", GetType(Integer))
        dtXmlEncargos.Columns.Add("SituacaoTributariaPISCOFINS", GetType(Integer))
        dtXmlEncargos.Columns.Add("SituacaoTributariaIPI", GetType(Integer))
        dtXmlEncargos.Columns.Add("EstadoOrigem", GetType(String))
        dtXmlEncargos.Columns.Add("EstadoDestino", GetType(String))
        dtXmlEncargos.Columns.Add("Base", GetType(Decimal))
        dtXmlEncargos.Columns.Add("Percentual", GetType(Decimal))
        dtXmlEncargos.Columns.Add("PercentualExibicao", GetType(Decimal))
        dtXmlEncargos.Columns.Add("Valor", GetType(Decimal))

        '--------ICMS-------------------------------------------------------

        Dim detTable As DataTable = dsXml.Tables("det")

        If detTable IsNot Nothing Then

            For Each detRow As DataRow In detTable.Rows

                Dim rowIndex As Integer = detTable.Rows.IndexOf(detRow)

                If rowIndex = posicaoGrid Then

                    ' Verifique se a tabela "imposto" está presente nos relacionamentos
                    Dim impostoRelations = detRow.Table.ChildRelations.Cast(Of DataRelation)().Where(Function(rel) rel.ChildTable.TableName = "imposto").ToList()

                    For Each impostoRelation As DataRelation In impostoRelations

                        Dim impostoRows As DataRow() = detRow.GetChildRows(impostoRelation)

                        For Each impostoRow As DataRow In impostoRows

                            ' Verifique os filhos da tabela "imposto" pelo nome
                            Dim impostoChildRelations = impostoRow.Table.ChildRelations.Cast(Of DataRelation)().Where(Function(rel) rel.ChildTable.TableName = "ICMS" OrElse rel.ChildTable.TableName = "IPI" OrElse rel.ChildTable.TableName = "PIS" OrElse rel.ChildTable.TableName = "COFINS").ToList()

                            For Each impostoChildRelation As DataRelation In impostoChildRelations

                                Dim impostoChildRows As DataRow() = impostoRow.GetChildRows(impostoChildRelation)

                                ' Itere sobre os dados da tabela "ICMS", "IPI", "PIS" ou "COFINS"
                                For Each impostoChildRow As DataRow In impostoChildRows

                                    Select Case impostoChildRow.Table.TableName

                                        Case "ICMS"

                                            '----------------------------ICMS00

                                            Dim icmsRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMS00"))

                                            If icmsRows.Count > 0 Then

                                                Dim impostoICMS As DataRow = dtXmlEncargos.NewRow()
                                                impostoICMS("Codigo") = "ICMS"

                                                For Each icmsRow As DataRow In icmsRows

                                                    If icmsRow.Table.Columns.Contains("CST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("CST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMS("SituacaoTributaria") = icmsRow("CST").ToString()
                                                        End If
                                                    End If

                                                    If icmsRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMS("Base") = icmsRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsRow.Table.Columns.Contains("pICMS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("pICMS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMS("Percentual") = icmsRow("pICMS").ToString().Replace(".", ",")
                                                            impostoICMS("PercentualExibicao") = icmsRow("pICMS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsRow.Table.Columns.Contains("vICMS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("vICMS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMS("Valor") = icmsRow("vICMS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoICMS("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoICMS("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoICMS)

                                            End If

                                            '----------------------------ICMS60

                                            Dim icmsRowsST As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMS60"))

                                            If icmsRowsST.Count > 0 Then

                                                Dim impostoICMSST As DataRow = dtXmlEncargos.NewRow()
                                                impostoICMSST("Codigo") = "ICMS-ST"

                                                For Each icmsRow As DataRow In icmsRowsST

                                                    If icmsRow.Table.Columns.Contains("CST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("CST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSST("SituacaoTributaria") = icmsRow("CST").ToString()
                                                        End If
                                                    End If

                                                    If icmsRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSST("Base") = icmsRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsRow.Table.Columns.Contains("pICMS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("pICMS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSST("Percentual") = icmsRow("pICMS").ToString().Replace(".", ",")
                                                            impostoICMSST("PercentualExibicao") = icmsRow("pICMS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsRow.Table.Columns.Contains("vICMS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsRow("vICMS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSST("Valor") = icmsRow("vICMS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoICMSST("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoICMSST("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoICMSST)

                                            End If

                                            '----------------------------ICMSUFDEST

                                            Dim icmsUFDestRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMSUFDEST"))

                                            If icmsUFDestRows.Count > 0 Then

                                                Dim impostoICMSUFDEST As DataRow = dtXmlEncargos.NewRow()
                                                impostoICMSUFDEST("Codigo") = "ICMS DIFERENCIAL"

                                                For Each icmsUFDestRow As DataRow In icmsUFDestRows

                                                    If icmsUFDestRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsUFDestRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSUFDEST("Base") = icmsUFDestRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsUFDestRow.Table.Columns.Contains("pICMSUFDEST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsUFDestRow("pICMSUFDEST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSUFDEST("Percentual") = icmsUFDestRow("pICMSUFDEST").ToString().Replace(".", ",")
                                                            impostoICMSUFDEST("PercentualExibicao") = icmsUFDestRow("pICMSUFDEST").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsUFDestRow.Table.Columns.Contains("vICMSUFDEST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsUFDestRow("vICMSUFDEST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSUFDEST("Valor") = icmsUFDestRow("vICMSUFDEST").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoICMSUFDEST("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoICMSUFDEST("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoICMSUFDEST)

                                            End If

                                            '----------------------------ICMSUFDEST

                                            Dim icmsdesonRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMSUFDEST"))

                                            If icmsdesonRows.Count > 0 Then

                                                Dim impostoICMSDESON As DataRow = dtXmlEncargos.NewRow()
                                                impostoICMSDESON("Codigo") = "ICMS DESONERADO"

                                                For Each icmsUFDestRow As DataRow In icmsdesonRows

                                                    If icmsUFDestRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsUFDestRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSDESON("Base") = icmsUFDestRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsUFDestRow.Table.Columns.Contains("pICMSUFDEST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsUFDestRow("pICMSDeson")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSDESON("Percentual") = icmsUFDestRow("pICMSDeson").ToString().Replace(".", ",")
                                                            impostoICMSDESON("PercentualExibicao") = icmsUFDestRow("pICMSDeson").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If icmsUFDestRow.Table.Columns.Contains("vICMSDeson") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(icmsUFDestRow("vICMSDeson")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoICMSDESON("Valor") = icmsUFDestRow("vICMSDeson").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoICMSDESON("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoICMSDESON("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoICMSDESON)

                                            End If

                                        Case "IPI"

                                            Dim ipiRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "IPINT"))

                                            If ipiRows.Count > 0 Then

                                                Dim impostoIPI As DataRow = dtXmlEncargos.NewRow()
                                                impostoIPI("Codigo") = "IPI"

                                                For Each ipiRow As DataRow In ipiRows

                                                    ' Verifica se a coluna "CST" existe na linha atual
                                                    If ipiRow.Table.Columns.Contains("CST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(ipiRow("CST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoIPI("SituacaoTributariaIPI") = ipiRow("CST").ToString()
                                                        End If
                                                    End If

                                                    If ipiRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(ipiRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoIPI("Base") = ipiRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If ipiRow.Table.Columns.Contains("pIPI") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(ipiRow("pIPI")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoIPI("Percentual") = ipiRow("pIPI").ToString().Replace(".", ",")
                                                            impostoIPI("PercentualExibicao") = ipiRow("pIPI").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If ipiRow.Table.Columns.Contains("vIPI") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(ipiRow("vIPI")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoIPI("Valor") = ipiRow("vIPI").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoIPI("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoIPI("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoIPI)

                                            End If

                                        Case "PIS"

                                            Dim pisRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "PISAliq"))

                                            If pisRows.Count > 0 Then

                                                Dim impostoPIS As DataRow = dtXmlEncargos.NewRow()
                                                impostoPIS("Codigo") = "PIS"

                                                For Each pisRow As DataRow In pisRows

                                                    ' Verifica se a coluna "CST" existe na linha atual
                                                    If pisRow.Table.Columns.Contains("CST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisRow("CST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPIS("SituacaoTributariaPISCOFINS") = pisRow("CST").ToString()
                                                        End If
                                                    End If

                                                    If pisRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPIS("Base") = pisRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If pisRow.Table.Columns.Contains("pPIS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisRow("pPIS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPIS("Percentual") = pisRow("pPIS").ToString().Replace(".", ",")
                                                            impostoPIS("PercentualExibicao") = pisRow("pPIS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If pisRow.Table.Columns.Contains("vPIS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisRow("vPIS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPIS("Valor") = pisRow("vPIS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoPIS("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoPIS("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoPIS)

                                            End If

                                        Case "COFINS"

                                            Dim cofinsRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "COFINSAliq"))

                                            If cofinsRows.Count > 0 Then

                                                Dim impostoCOFINS As DataRow = dtXmlEncargos.NewRow()
                                                impostoCOFINS("Codigo") = "COFINS"

                                                For Each cofinsRow As DataRow In cofinsRows

                                                    ' Verifica se a coluna "CST" existe na linha atual
                                                    If cofinsRow.Table.Columns.Contains("CST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsRow("CST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINS("SituacaoTributariaPISCOFINS") = cofinsRow("CST").ToString()
                                                        End If
                                                    End If

                                                    If cofinsRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINS("Base") = cofinsRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If cofinsRow.Table.Columns.Contains("pCOFINS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsRow("pCOFINS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINS("Percentual") = cofinsRow("pCOFINS").ToString().Replace(".", ",")
                                                            impostoCOFINS("PercentualExibicao") = cofinsRow("pCOFINS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If cofinsRow.Table.Columns.Contains("vCOFINS") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsRow("vCOFINS")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINS("Valor") = cofinsRow("vCOFINS").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoCOFINS("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoCOFINS("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoCOFINS)

                                            End If

                                        Case "PISST"

                                            Dim pisstRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "PISST"))

                                            If pisstRows.Count > 0 Then

                                                Dim impostoPISST As DataRow = dtXmlEncargos.NewRow()
                                                impostoPISST("Codigo") = "PISST"

                                                For Each pisstRow As DataRow In pisstRows

                                                    If pisstRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisstRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPISST("Base") = pisstRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If pisstRow.Table.Columns.Contains("pPISST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisstRow("pPISST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPISST("Percentual") = pisstRow("pPISST").ToString().Replace(".", ",")
                                                            impostoPISST("PercentualExibicao") = pisstRow("pPISST").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If pisstRow.Table.Columns.Contains("vPISST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(pisstRow("vPISST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoPISST("Valor") = pisstRow("vPISST").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoPISST("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoPISST("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoPISST)

                                            End If

                                        Case "COFINSST"

                                            Dim cofinsstRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "COFINSST"))

                                            If cofinsstRows.Count > 0 Then

                                                Dim impostoCOFINSST As DataRow = dtXmlEncargos.NewRow()
                                                impostoCOFINSST("Codigo") = "COFINSST"

                                                For Each cofinsstRow As DataRow In cofinsstRows

                                                    If cofinsstRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsstRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINSST("Base") = cofinsstRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If cofinsstRow.Table.Columns.Contains("pCOFINSST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsstRow("pCOFINSST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINSST("Percentual") = cofinsstRow("pCOFINSST").ToString().Replace(".", ",")
                                                            impostoCOFINSST("PercentualExibicao") = cofinsstRow("pCOFINSST").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If cofinsstRow.Table.Columns.Contains("vCOFINSST") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(cofinsstRow("vCOFINSST")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoCOFINSST("Valor") = cofinsstRow("vCOFINSST").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoCOFINSST("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoCOFINSST("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoCOFINSST)

                                            End If

                                        Case "ISSQN"

                                            Dim issqnRows As DataRow() = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ISSQN"))

                                            If issqnRows.Count > 0 Then

                                                Dim impostoISSQN As DataRow = dtXmlEncargos.NewRow()
                                                impostoISSQN("Codigo") = "ISSQN"

                                                For Each issqnRow As DataRow In issqnRows

                                                    If issqnRow.Table.Columns.Contains("vBC") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(issqnRow("vBC")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoISSQN("Base") = issqnRow("vBC").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If issqnRow.Table.Columns.Contains("pISSQN") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(issqnRow("pISSQN")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoISSQN("Percentual") = issqnRow("pISSQN").ToString().Replace(".", ",")
                                                            impostoISSQN("PercentualExibicao") = issqnRow("pISSQN").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                    If issqnRow.Table.Columns.Contains("vISSQN") Then
                                                        ' Verifica se a coluna "CST" não está vazia
                                                        If Not IsDBNull(issqnRow("vISSQN")) Then
                                                            ' Obtém o valor de "CST" e faz algo com ele
                                                            impostoISSQN("Valor") = issqnRow("vISSQN").ToString().Replace(".", ",")
                                                        End If
                                                    End If

                                                Next

                                                impostoISSQN("EstadoOrigem") = dsXml.Tables("EnderEmit").Rows(0).Item("UF")
                                                impostoISSQN("EstadoDestino") = dsXml.Tables("EnderDest").Rows(0).Item("UF")

                                                dtXmlEncargos.Rows.Add(impostoISSQN)

                                            End If

                                        Case "ICMSUFDEST"



                                    End Select
                                Next
                            Next
                        Next
                    Next

                End If

            Next

        End If

        DgEncargosXML.DataSource = dtXmlEncargos
        DgEncargosXML.DataBind()


    End Sub

    Function IndiceTabela(tabela As DataSet, nomeTabela As String) As Integer

        For i As Integer = 0 To tabela.Tables.Count - 1

            If tabela.Tables(i).TableName = nomeTabela Then

                Return i

            End If

        Next

        Return -1

    End Function

    Private Sub AtualizarEncargos(ByVal index As Integer)

        If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES AndAlso objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            MsgBox(Me.Page, "Encargos não podem ser ajustados por aqui, caso seja necessário ajuste o pedido e refaça a nota fiscal.")
            Exit Sub
        End If

        Dim Base As Decimal = CType(DgEncargos.Rows(index).FindControl("txtBaseEncargoItem"), TextBox).Text
        Dim Percentual As Decimal = CType(DgEncargos.Rows(index).FindControl("txtPercentualEncargoItem"), TextBox).Text
        Dim Valor As Decimal = CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text

        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Sinal = "-" Then

            Dim encargoLiquido As ListNotaFiscalXItemXEncargo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos().Clone()

            For Each encargo As NotaFiscalXItemXEncargo In encargoLiquido
                If encargo.Codigo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo Then
                    encargo.Valor = Valor
                    Exit For
                End If
            Next

            If encargoLiquido.Where(Function(x) x.Codigo = "LIQUIDO").Count() > 0 AndAlso encargoLiquido.Where(Function(x) x.Codigo = "LIQUIDO").FirstOrDefault().Valor < 0 Then
                MsgBox(Me.Page, "O valor liquido do item não pode ser negativo. Qualquer dúvida entre em contato com o Suporte.")
                Exit Sub
            End If

        End If

            If objNotaFiscal.NFG Then
            'Notas Fiscais Gerais
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FACS" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "SENAR" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL JUDICIAL" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL" Then

                Dim vlrOriginal As Decimal = 0
                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                    vlrOriginal = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base / 1000 * objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual, 2, MidpointRounding.AwayFromZero)
                Else
                    vlrOriginal = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base * objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual / 100, 2, MidpointRounding.AwayFromZero)
                End If
                If Valor > (vlrOriginal + 1) OrElse Valor < (vlrOriginal - 1) Then
                    Valor = vlrOriginal
                    CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = Valor
                    MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                    Exit Sub
                End If
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "DESCONTOS" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "DESP. ACESSORIA" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True

            Dim valorApuracao As Decimal = Math.Round(Base * Percentual / 100, 2, MidpointRounding.AwayFromZero)

            If Percentual = 0 Then valorApuracao = Valor

            If Valor > (valorApuracao + 1) OrElse
                Valor < (valorApuracao - 1) Then
                MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                Exit Sub
            Else
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base = Base
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = Percentual
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).PercentualExibicao = Percentual
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = Valor
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FACS" And Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" And Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" Then
                    objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = Math.Round((objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor * 1000) / objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base, 6, MidpointRounding.AwayFromZero)
                End If
            Else
                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base = 0 OrElse
                    objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = 0 OrElse
                    objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = 0 Then

                    If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo.Contains("FGTS") AndAlso
                        Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo.Contains("FRETES") AndAlso
                        Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo.Contains("DESCONTOS") AndAlso
                        Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo.Contains("DESP. ACESSORIA") AndAlso
                        Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo.Contains("ICMS MONO") AndAlso
                        Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo.Contains("INSS") Then
                        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base = 0
                        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = 0
                        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = 0
                    End If
                End If
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "DESCONTOS" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "DESP. ACESSORIA" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False

            If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).IUD = "I" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).IUD = "U"

            DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
            DgEncargos.DataBind()
            objNotaFiscal.AtualizaTotais()
            HabilitarCampos()

        Else

            'Notas Fiscais
            Dim objEncargoXTaxa As New EncargoXTaxa()

            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FACS" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FETHAB GADO" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FABOV" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "SENAR" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDEMS" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL JUDICIAL" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL" Then

                If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDEMS" Then
                    objEncargoXTaxa.SelecionarVigente(objNotaFiscal.Empresa.CodigoEstado, "UFERMS", objNotaFiscal.Movimento, objNotaFiscal.Itens(hdfIndex.Value).CodigoProduto)
                End If

                Dim vlrOriginal As Decimal = 0

                If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                    If Not objEncargoXTaxa.Estado Is Nothing AndAlso objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDEMS" Then
                        vlrOriginal = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual) / 100) * (objNotaFiscal.Itens(hdfIndex.Value).QuantidadeFiscal / 1000), 2, MidpointRounding.AwayFromZero)
                    Else
                        vlrOriginal = Math.Round(objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base / 1000 * objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual, 2, MidpointRounding.AwayFromZero)
                    End If
                Else
                    vlrOriginal = Math.Round(objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base * objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual / 100, 2, MidpointRounding.AwayFromZero)
                End If

                If Valor > (vlrOriginal + 1) OrElse Valor < (vlrOriginal - 1) Then
                    Valor = vlrOriginal
                    CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = Valor
                    MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00")
                    Exit Sub
                End If
            End If

            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "ICMS-ST" Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "IPI" Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo.Contains("DESPESAS") Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo.Contains("DESCONTOS") Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo.Contains("DESP. ACESSORIA") Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True

            Dim valorApuracao As Decimal = 0
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then

                If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDEMS" Then

                    objEncargoXTaxa.SelecionarVigente(objNotaFiscal.Empresa.CodigoEstado, "UFERMS", objNotaFiscal.Movimento, objNotaFiscal.Itens(hdfIndex.Value).CodigoProduto)

                    If Not objEncargoXTaxa.Estado Is Nothing Then
                        valorApuracao = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual) / 100) * (objNotaFiscal.Itens(hdfIndex.Value).QuantidadeFiscal / 1000), 2, MidpointRounding.AwayFromZero)
                    Else
                        valorApuracao = 0
                    End If
                Else
                    valorApuracao = Math.Round(Base * Percentual / 1000, 2, MidpointRounding.AwayFromZero)
                End If
            Else
                valorApuracao = Math.Round(Base * Percentual / 100, 2, MidpointRounding.AwayFromZero)
            End If

            If Percentual > 0 AndAlso (Valor > (valorApuracao + 1) OrElse
                Valor < (valorApuracao - 1)) Then
                CType(DgEncargos.Rows(index).FindControl("txtBaseEncargoItem"), TextBox).Text = objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base
                CType(DgEncargos.Rows(index).FindControl("txtPercentualEncargoItem"), TextBox).Text = objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual
                CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Valor
                MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                Exit Sub
            Else
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base = Base
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual = Percentual
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).PercentualExibicao = Percentual
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Valor = Valor
            End If

            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                If Not objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FACS" And
                    Not objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" And
                    Not objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" And
                    Not objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDERSUL" And
                    Not objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDEMS" Then
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual = Math.Round((objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Valor * 1000) / objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base, 6, MidpointRounding.AwayFromZero)
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).PercentualExibicao = objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual
                End If
            Else
                If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "ICMS-ST" OrElse
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "IPI" OrElse
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FRETES" OrElse
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "DESCONTOS" OrElse
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "DESP. ACESSORIA" OrElse
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "DESPESAS" Then
                    'NÃO ZERAR
                ElseIf objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base = 0 Then
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base = 0
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual = 0
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).PercentualExibicao = 0
                    objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Valor = 0
                End If
            End If

            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "ICMS-ST" Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "IPI" Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo.Contains("DESPESAS") Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo.Contains("DESCONTOS") Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False
            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo.Contains("DESP. ACESSORIA") Then objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False

            If objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
                OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("TRANSFERENCIA DE CREDITO DE ICMS") Then
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Base = 0
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Percentual = 0
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).PercentualExibicao = 0
                objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Valor = CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text
            End If

            DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
            DgEncargos.DataBind()
            objNotaFiscal.AtualizaTotais()
            HabilitarCampos()

            If objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FACS" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FETHAB GADO" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "SENAR" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FABOV" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNDEMS" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL JUDICIAL" OrElse
               objNotaFiscal.Itens(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL" Then

                For Each enc In objNotaFiscal.Itens(0).Encargos
                    If enc.Codigo.Contains("FUNRURAL") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNRURAL - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FABOV") Then
                        observacaoValores.Value = observacaoValores.Value & "|FABOV - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("SENAR") Then
                        observacaoValores.Value = observacaoValores.Value & "|SENAR - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FETHAB") Then
                        observacaoValores.Value = observacaoValores.Value & "|FETHAB - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("IAGRO") Then
                        observacaoValores.Value = observacaoValores.Value & "|IAGRO - " & enc.Percentual.ToString.Replace(",", ".") & IIf(enc.Encargo.Operador = "*", " por Unidade - R$ ", " % - R$ ") & "= " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FACS") Then
                        observacaoValores.Value = observacaoValores.Value & "|FACS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FUNDERSUL") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNDERSUL - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    ElseIf enc.Codigo.Contains("FUNDEMS") Then
                        observacaoValores.Value = observacaoValores.Value & "|FUNDEMS - " & enc.Percentual.ToString.Replace(",", ".") & " % - R$ " & enc.Valor.ToString.Replace(".", "").Replace(",", ".")
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub HabilitarCampos()
        Dim rowIndex As Integer
        For rowIndex = 0 To DgEncargos.Rows.Count - 1

            If objNotaFiscal.NFG Then
                'Notas Fiscais Gerais
                If DgEncargos.Rows(rowIndex).Cells(0).Text() = "LIQUIDO" OrElse
                   DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("PRODUTO") Then
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = False
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = False
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = False
                    CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = False
                Else
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                    CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                End If
            Else
                'Notas Fiscais

                If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES AndAlso objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = False
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = False
                    CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = False

                ElseIf Not objNotaFiscal.IUD = "I" AndAlso objNotaFiscal.Eletronica AndAlso objNotaFiscal.NossaEmissao Then
                    If DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("ICMS") OrElse
                        DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("ISS") OrElse
                        DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("IPI") OrElse
                        DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("PIS") OrElse
                        DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("COFINS") Then
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                    End If
                Else
                    If DgEncargos.Rows(rowIndex).Cells(0).Text() = "FRETES" OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text() = "DESCONTOS" OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text() = "SEGURO" OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FUNRURAL") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("SENAR") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FACS") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FETHAB") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("IAGRO") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FETHAB GADO") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FABOV") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FUNDERSUL") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FUNDEMS") OrElse
                       DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("ADUANEIRAS") Then
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                    ElseIf DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("ICMS") OrElse
                           DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("IPI") OrElse
                           DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("ISS") OrElse
                           DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("PIS") OrElse
                           DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("COFINS") OrElse
                           DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("DESPESAS") OrElse
                           DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("DESP. ACESSORIA") Then
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                    ElseIf DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("FUNDO FECP") AndAlso
                        objNotaFiscal.SubOperacao.FinalidadeDaNota = 2 AndAlso
                        objNotaFiscal.Itens(hdfIndex.Value).ValorTotal = 0 Then
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = False
                        CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                        CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                    End If
                End If
            End If
        Next

    End Sub
#End Region

End Class