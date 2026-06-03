Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AlterarEncargoCTE
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AlterarEncargo", "ACESSAR") Then
                VerificaUnidade()
                Limpar()
            Else
                MsgBox(Me.Page, "Página temporariamente indisponível!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value.ToString) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub VerificaUnidade()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa, True)
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue)
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClienteALTxENC" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objClienteALTxENC" & HID.Value.ToString) Is Nothing Then
                Dim cli As [Lib].Negocio.Cliente = Session("objClienteALTxENC" & HID.Value.ToString)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteALTxENC" & HID.Value.ToString)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        Session.Remove("objClienteALTxENC" & HID.Value)
        Session.Remove("ssAlterarEncargo" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty

        idGridItem.Value = -1

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        idBeneficio.Visible = False
        idVigencia.Visible = False
        idVigenciaNova.Visible = False
        idVersao.Visible = False
        idVersaoNova.Visible = False
        tabEncargos.Visible = False
        tabEncargosNovos.Visible = False
        idSubtitulo.Visible = False

        LiberaEmpresa()

        gridItens.DataSource = Nothing
        gridItens.DataBind()

        HID.Value = Guid.NewGuid().ToString
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Sub CarregarItem()
        Try
            If Not Session("objNFConsultaNXI" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaNXI" & HID.Value), NotaFiscal))

                If Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
                    MsgBox(Me.Page, "Movimento da Nota Fiscal já Fechado para esta data.")
                    Exit Sub
                ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, objNotaFiscal.Movimento, "CONTABIL") Then
                    MsgBox(Me.Page, "Movimento Contábil já Fechado para esta data.")
                    Exit Sub
                ElseIf Not objNotaFiscal.TipoDeDocumento.Codigo = 57 Then
                    MsgBox(Me.Page, "O documento não é um CTE!.")
                    Exit Sub
                End If

                txtES.Text = objNotaFiscal.EntradaSaida.ToString()
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie

                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False

                gridItens.DataSource = objNotaFiscal.Itens
                gridItens.DataBind()

                SalvaNotaFiscal()

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridItens_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridItens.SelectedIndexChanged
        Try
            RecuperaNotaFiscal()

            idVigencia.Visible = True
            idVigenciaNova.Visible = True
            idVersao.Visible = True
            idVersaoNova.Visible = True
            tabEncargos.Visible = True
            tabEncargosNovos.Visible = True
            idSubtitulo.Visible = True

            lblProduto.Text = gridItens.SelectedRow.Cells(2).Text() & " - " & gridItens.SelectedRow.Cells(3).Text()
            idGridItem.Value = gridItens.SelectedIndex

            hdCentroDeCusto.Value = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos.EncProduto.CentroDeCusto

            gridEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos
            gridEncargos.DataBind()

            gridEncargosNovos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos
            gridEncargosNovos.DataBind()

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.CodigoBeneficio.Length > 0 Then
                idBeneficio.Visible = True
                idBeneficioNovo.Visible = True
                txtBeneficioICMS.Text = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.BeneficioICMS.Codigo & " - " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.BeneficioICMS.Descricao
                txtBeneficioICMSNovo.Text = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.BeneficioICMS.Codigo & " - " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.BeneficioICMS.Descricao
            Else
                idBeneficio.Visible = False
                idBeneficioNovo.Visible = False
                txtBeneficioICMS.Text = String.Empty
                txtBeneficioICMSNovo.Text = String.Empty
            End If

            Dim Parametros As New OperacaoXEstado

            If objNotaFiscal.IUD = "U" Then
                Parametros.Codigo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacaoEstado
            ElseIf Not objNotaFiscal.NossaEmissao Then
                Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
                Parametros.CodigoGrupoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Produto.CodigoGrupo
                Parametros.CodigoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoProduto
                Parametros.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacao
                Parametros.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoSubOperacao

                Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado

                Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado
                If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then
                    Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
                End If

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                    If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                        objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                        objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                        Parametros.EstadoDestino = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado
                        If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
                    End If
                End If
            Else
                Parametros.Codigo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacaoEstado
            End If

            Dim ObjListVersoes As New ListOperacaoXEstado(Parametros)

            ddlVersao.Items.Clear()

            ddlVersao.DataTextField = "DESCRICAO"
            ddlVersao.DataValueField = "CODIGO"

            ddlVersao.DataSource = (From x In ObjListVersoes
                                    Order By x.InicioVigencia, x.Codigo
                                    Select x.Codigo, Descricao = x.InicioVigencia.ToString("dd-MM-yyyy") & "   ID-" & x.Codigo & "   " & x.UsuarioInclusao & "   " & x.DataHoraInclusao.ToString("dd-MM-yyyy") & IIf(x.Ativo, "   Ativo", "   Cancelada"))
            ddlVersao.DataBind()

            ddlVersao.SelectedIndex = 0

            With ddlVersao
                ddlVersao.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacaoEstado))
            End With

            Parametros = New OperacaoXEstado

            If Not objNotaFiscal.NossaEmissao Then
                Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
                Parametros.CodigoGrupoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Produto.CodigoGrupo
                Parametros.CodigoProduto = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoProduto
                Parametros.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacao
                'Parametros.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoSubOperacao

                Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado

                Parametros.EstadoDestino = objNotaFiscal.Cliente.CodigoEstado
                If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then
                    Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
                End If

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                    If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                        objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                        objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                        Parametros.EstadoDestino = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado
                        If Not objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then Parametros.RegiaoDestino = objNotaFiscal.Cliente.Estado.Regiao
                    End If
                End If
            Else
                Parametros.Codigo = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacaoEstado
            End If

            ObjListVersoes = New ListOperacaoXEstado(Parametros)

            ddlVersaoNova.Items.Clear()

            ddlVersaoNova.DataTextField = "DESCRICAO"
            ddlVersaoNova.DataValueField = "CODIGO"

            ddlVersaoNova.DataSource = (From x In ObjListVersoes
                                        Order By x.InicioVigencia, x.Codigo
                                        Select x.Codigo, Descricao = x.InicioVigencia.ToString("dd-MM-yyyy") & "   ID-" & x.Codigo & "   " & x.UsuarioInclusao & "   " & x.DataHoraInclusao.ToString("dd-MM-yyyy") & IIf(x.Ativo, "   Ativo", "   Cancelada"))
            ddlVersaoNova.DataBind()

            ddlVersaoNova.SelectedIndex = 0

            With ddlVersaoNova
                ddlVersaoNova.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacaoEstado))
            End With

            ddlVersaoNova.Enabled = True

            HabilitarCampos()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub HabilitarCampos()
        Dim rowIndex As Integer
        For rowIndex = 0 To gridEncargosNovos.Rows.Count - 1

            If objNotaFiscal.NFG Then
                'Notas Fiscais Gerais
                If gridEncargosNovos.Rows(rowIndex).Cells(0).Text() = "LIQUIDO" OrElse
                   gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("PRODUTO") Then
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = False
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = False
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = False
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = False
                Else
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                End If
            Else
                'Notas Fiscais
                If objNotaFiscal.IUD = "U" Then
                    If objNotaFiscal.NossaEmissao Then
                        If gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("ISS") OrElse
                            gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("IPI") OrElse
                            gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("PIS") OrElse
                            gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("COFINS") Then
                            'CType(gridEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                            'CType(gridEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                            'CType(gridEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                            'CType(gridEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                        End If
                        If gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("DESCONTOS") Then
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                        End If
                    Else
                        If gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("ICMS") OrElse
                               gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("IPI") OrElse
                               gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("ISS") OrElse
                               gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("PIS") OrElse
                               gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("COFINS") OrElse
                               gridEncargosNovos.Rows(rowIndex).Cells(0).Text().Contains("DESP. ACESSORIA") Then
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                            CType(gridEncargosNovos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
                        End If
                    End If
                Else
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = False
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = False
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = False
                    CType(gridEncargosNovos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = False
                End If
            End If
        Next
    End Sub

    Protected Sub gridEncargosNovos_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridEncargosNovos.RowCreated
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim btnEncargoItem As Button = CType(e.Row.FindControl("btnEncargoItem"), Button)
            btnEncargoItem.CommandArgument = e.Row.RowIndex.ToString()
        End If
    End Sub

    Protected Sub gridEncargosNovos_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gridEncargosNovos.RowCommand
        If e.CommandName = "OK" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            AtualizarEncargos(index)
            lnkAtualizar.Parent.Visible = True
        End If
    End Sub

    Protected Sub btnRecarregar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRecarregar.Click
        Try
            RecuperaNotaFiscal()

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

            For Each item In objNotaFiscal.Itens
                item.IUD = "U"
            Next

            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacaoEstado = ddlVersaoNova.SelectedValue

            objNotaFiscal.CarregandoNota = True
            Dim nXe = New ListNotaFiscalXItemXEncargo(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})

            If nXe.Count > 0 Then
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos = Nothing
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos = nXe

                objNotaFiscal.CodigoOperacao = nXe.EncProduto.CodigoOperacao
                objNotaFiscal.CodigoSubOperacao = nXe.EncProduto.CodigoSubOperacao

                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoOperacao = nXe.EncProduto.CodigoOperacao
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).CodigoSubOperacao = nXe.EncProduto.CodigoSubOperacao

                If hdCentroDeCusto.Value > 0 Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos.EncProduto.CentroDeCusto = hdCentroDeCusto.Value

                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.CodigoBeneficio.Length > 0 Then
                    idBeneficioNovo.Visible = True
                    txtBeneficioICMSNovo.Text = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.BeneficioICMS.Codigo & " - " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).OperacaoEstado.BeneficioICMS.Descricao
                Else
                    idBeneficioNovo.Visible = False
                    txtBeneficioICMSNovo.Text = String.Empty
                End If

                gridEncargosNovos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos
                gridEncargosNovos.DataBind()

                HabilitarCampos()

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                    objNotaFiscal.TipoDeDocumentoFrete = Nothing

                    objNotaFiscal.CodigoCliente = fCliente
                    objNotaFiscal.EnderecoCliente = fEndCliente
                End If

                SalvaNotaFiscal()

                lnkAtualizar.Parent.Visible = True

            Else
                MsgBox(Me.Page, "Encargos não foram encotrados/recarregados")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarEncargos(ByVal index As Integer)

        RecuperaNotaFiscal()

        Dim Base As Decimal = CType(gridEncargosNovos.Rows(index).FindControl("txtBaseEncargoItem"), TextBox).Text
        Dim Percentual As Decimal = CType(gridEncargosNovos.Rows(index).FindControl("txtPercentualEncargoItem"), TextBox).Text
        Dim Valor As Decimal = CType(gridEncargosNovos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text

        If objNotaFiscal.NFG Then
            'Notas Fiscais Gerais
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FACS" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FETHAB" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "SENAR" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FUNRURAL JUDICIAL" OrElse
               objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FUNRURAL" Then

                Dim vlrOriginal As Decimal = 0
                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                    vlrOriginal = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Base / 1000 * objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Percentual, 2, MidpointRounding.AwayFromZero)
                Else
                    vlrOriginal = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Base * objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Percentual / 100, 2, MidpointRounding.AwayFromZero)
                End If
                If Valor > (vlrOriginal + 1) OrElse Valor < (vlrOriginal - 1) Then
                    Valor = vlrOriginal
                    CType(gridEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = Valor
                    MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                    Exit Sub
                End If
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Encargo.Atualizacao = True

            Dim valorApuracao As Decimal = Math.Round(Base * Percentual / 100, 2, MidpointRounding.AwayFromZero)

            If Percentual = 0 Then valorApuracao = Valor

            If Valor > (valorApuracao + 1) OrElse
                Valor < (valorApuracao - 1) Then
                MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                Exit Sub
            Else
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Base = Base
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Percentual = Percentual
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).PercentualExibicao = Percentual
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Valor = Valor
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FACS" And Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FETHAB" Then
                    objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Percentual = Math.Round((objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Valor * 1000) / objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Base, 6, MidpointRounding.AwayFromZero)
                End If
            Else
                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Base = 0 OrElse
                    objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Percentual = 0 OrElse
                    objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Valor = 0 Then

                    If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo.Contains("FGTS") AndAlso
                        Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo.Contains("INSS") Then
                        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Base = 0
                        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Percentual = 0
                        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Valor = 0
                    End If
                End If
            End If

            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos(index).Encargo.Atualizacao = False

            If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).IUD = "I" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).IUD = "U"

            gridEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos
            gridEncargos.DataBind()
            objNotaFiscal.AtualizaTotais()
            HabilitarCampos()
        Else

            'Notas Fiscais
            Dim objEncargoXTaxa As New EncargoXTaxa()

            If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FACS" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FETHAB" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FETHAB GADO" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FABOV" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "SENAR" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDEMS" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNRURAL JUDICIAL" OrElse
               objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNRURAL" Then

                If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDEMS" Then
                    objEncargoXTaxa.SelecionarVigente(objNotaFiscal.Empresa.CodigoEstado, "UFERMS", objNotaFiscal.Movimento, objNotaFiscal.Itens(idGridItem.Value).CodigoProduto)
                End If

                Dim vlrOriginal As Decimal = 0

                If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                    If Not objEncargoXTaxa.Estado Is Nothing AndAlso objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDEMS" Then
                        vlrOriginal = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual) / 100) * (objNotaFiscal.Itens(idGridItem.Value).QuantidadeFiscal / 1000), 2, MidpointRounding.AwayFromZero)
                    Else
                        vlrOriginal = Math.Round(objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base / 1000 * objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual, 2, MidpointRounding.AwayFromZero)
                    End If
                Else
                    vlrOriginal = Math.Round(objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base * objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual / 100, 2, MidpointRounding.AwayFromZero)
                End If

                If Valor > (vlrOriginal + 1) OrElse Valor < (vlrOriginal - 1) Then
                    Valor = vlrOriginal
                    CType(gridEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = Valor
                    MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00")
                    Exit Sub
                End If
            End If

            If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Encargo.Atualizacao = True

            Dim valorApuracao As Decimal = 0
            If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then

                If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDERSUL" OrElse objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDEMS" Then

                    objEncargoXTaxa.SelecionarVigente(objNotaFiscal.Empresa.CodigoEstado, "UFERMS", objNotaFiscal.Movimento, objNotaFiscal.Itens(idGridItem.Value).CodigoProduto)

                    If Not objEncargoXTaxa.Estado Is Nothing Then
                        valorApuracao = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual) / 100) * (objNotaFiscal.Itens(idGridItem.Value).QuantidadeFiscal / 1000), 2, MidpointRounding.AwayFromZero)
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
                CType(gridEncargosNovos.Rows(index).FindControl("txtBaseEncargoItem"), TextBox).Text = objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base
                CType(gridEncargosNovos.Rows(index).FindControl("txtPercentualEncargoItem"), TextBox).Text = objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual
                CType(gridEncargosNovos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Valor
                MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                Exit Sub
            Else
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base = Base
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual = Percentual
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).PercentualExibicao = Percentual
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Valor = Valor
            End If

            If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                If Not objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FACS" And
                    Not objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FETHAB" And
                    Not objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDERSUL" And
                    Not objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FUNDEMS" Then
                    objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual = Math.Round((objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Valor * 1000) / objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base, 6, MidpointRounding.AwayFromZero)
                    objNotaFiscal.Itens(idGridItem.Value).Encargos(index).PercentualExibicao = objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual
                End If
            Else
                If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base = 0 Then
                    objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base = 0
                    objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual = 0
                    objNotaFiscal.Itens(idGridItem.Value).Encargos(index).PercentualExibicao = 0
                    objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Valor = 0
                End If
            End If

            If objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Encargo.Atualizacao = False

            If objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("TRANSFERENCIA DE CREDITO DE ICMS") Then
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Base = 0
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Percentual = 0
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).PercentualExibicao = 0
                objNotaFiscal.Itens(idGridItem.Value).Encargos(index).Valor = CType(gridEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text
            End If

            gridEncargosNovos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(idGridItem.Value).Encargos
            gridEncargosNovos.DataBind()

            objNotaFiscal.AtualizaTotais()

            HabilitarCampos()

            SalvaNotaFiscal()

        End If
    End Sub


    'Function validarCampos() As Boolean
    '    If chkIcms.Checked Then
    '        If String.IsNullOrWhiteSpace(cstIcms.Text) Then
    '            MsgBox(Me.Page, "Situação Tributária ICMS deve ser Informada")
    '            Return False
    '        End If
    '    End If

    '    If chkPis.Checked Then
    '        If cstPis.Text <> cstCofins.Text Or String.IsNullOrWhiteSpace(cstPis.Text) Then
    '            MsgBox(Me.Page, "Situação Tributário PIS não pode ser diferente do Cofins")
    '            Return False
    '        ElseIf Not String.IsNullOrWhiteSpace(BasePis.Text) AndAlso String.IsNullOrWhiteSpace(txtObsCofins.Text) Then
    '            MsgBox(Me.Page, "Observação PIS/COFINS deve ser informada")
    '            Return False
    '        End If
    '    End If

    '    If chkCofins.Checked Then
    '        If cstPis.Text <> cstCofins.Text Or String.IsNullOrWhiteSpace(cstCofins.Text) Then
    '            MsgBox(Me.Page, "Situação Tributário Cofins não pode ser diferente do Pis")
    '            Return False
    '        ElseIf Not String.IsNullOrWhiteSpace(BaseCofins.Text) AndAlso String.IsNullOrWhiteSpace(txtObsCofins.Text) Then
    '            MsgBox(Me.Page, "Observação PIS/COFINS deve ser informada")
    '            Return False
    '        End If
    '    End If

    '    If chkIPI.Checked Then
    '        If String.IsNullOrWhiteSpace(cstIPI.Text) Then
    '            MsgBox(Me.Page, "Situação Tributária IPI deve ser Informada")
    '            Return False
    '        End If
    '    End If

    '    If chkPisR.Checked Then
    '        If cstPisR.Text <> cstCofinsR.Text Or String.IsNullOrWhiteSpace(cstPisR.Text) Then
    '            MsgBox(Me.Page, "Situação Tributário PIS Retido não pode ser diferente do Cofins Retido")
    '            Return False
    '        ElseIf Not String.IsNullOrWhiteSpace(BasePisR.Text) AndAlso String.IsNullOrWhiteSpace(txtObsCofinsR.Text) Then
    '            MsgBox(Me.Page, "Observação PIS/COFINS Retido deve ser informada")
    '            Return False
    '        End If
    '    End If

    '    If chkCofinsR.Checked Then
    '        If cstPisR.Text <> cstCofinsR.Text Or String.IsNullOrWhiteSpace(cstCofinsR.Text) Then
    '            MsgBox(Me.Page, "Situação Tributário Cofins Retido não pode ser diferente do Pis Retido")
    '            Return False
    '        ElseIf Not String.IsNullOrWhiteSpace(BaseCofinsR.Text) AndAlso String.IsNullOrWhiteSpace(txtObsCofinsR.Text) Then
    '            MsgBox(Me.Page, "Observação PIS/COFINS deve ser informada")
    '            Return False
    '        End If
    '    End If

    '    If chkCSLLR.Checked Then
    '        If String.IsNullOrWhiteSpace(cstCSLLR.Text) Then
    '            MsgBox(Me.Page, "Situação Tributária CSLL deve ser Informada")
    '            Return False
    '        End If
    '    End If

    '    If chkPisCREPRE.Checked Then
    '        If cstPisCREPRE.Text <> cstCofinsCREPRE.Text Or String.IsNullOrWhiteSpace(cstPisCREPRE.Text) Then
    '            MsgBox(Me.Page, "Situação Tributário PIS Crédito Presumido não pode ser diferente do Cofins Crédito Presumido")
    '            Return False
    '        ElseIf Not String.IsNullOrWhiteSpace(BasePisCREPRE.Text) AndAlso String.IsNullOrWhiteSpace(txtObsCofinsCREPRE.Text) Then
    '            MsgBox(Me.Page, "Observação PIS/COFINS Crédito Presumido deve ser informada")
    '            Return False
    '        End If
    '    End If

    '    If chkCofinsCREPRE.Checked Then
    '        If cstPisCREPRE.Text <> cstCofinsCREPRE.Text Or String.IsNullOrWhiteSpace(cstCofinsCREPRE.Text) Then
    '            MsgBox(Me.Page, "Situação Tributário Cofins Crédito Presumido não pode ser diferente do Pis Crédito Presumido")
    '            Return False
    '        ElseIf Not String.IsNullOrWhiteSpace(BaseCofinsCREPRE.Text) AndAlso String.IsNullOrWhiteSpace(txtObsCofinsCREPRE.Text) Then
    '            MsgBox(Me.Page, "Observação PIS/COFINS Crédito Presumido deve ser informada")
    '            Return False
    '        End If
    '    End If

    '    Return True
    'End Function

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("NotaFiscalXItens", "ALTERAR") Then
                If Not String.IsNullOrWhiteSpace(txtES.Text) AndAlso Not String.IsNullOrWhiteSpace(txtNota.Text) AndAlso Not String.IsNullOrWhiteSpace(txtSerie.Text) Then
                    IniciarAlteracaoNotaFiscal()
                Else
                    MsgBox(Me.Page, "E/S, Número ou Série da Nota Fiscal não estão em conformidade, verifique.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
    '    Try
    '        If validarCampos() Then
    '            RecuperaNotaFiscal()

    '            Dim Sqls As New ArrayList

    '            'If (HttpContext.Current.Session("ssNomeUsuario") = "YURI" Or HttpContext.Current.Session("ssNomeUsuario") = "CONTABILINSOL02") AndAlso objNotaFiscal.Movimento.Year <= 2014 Then
    '            'Else
    '            '    If objNotaFiscal.NFG Then
    '            '        Dim cLote As Integer = 0
    '            '        If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or _
    '            '            objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or _
    '            '            objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
    '            '            cLote = 21
    '            '        Else
    '            '            cLote = 9
    '            '        End If
    '            '        objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, cLote)
    '            '    Else
    '            '        If objNotaFiscal.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CTRC) _
    '            '                AndAlso objNotaFiscal.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ReciboDeFrete) _
    '            '                AndAlso objNotaFiscal.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) _
    '            '                AndAlso objNotaFiscal.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Anulacao) Then
    '            '            objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls)
    '            '        Else
    '            '            objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
    '            '        End If
    '            '    End If
    '            'End If
    '            'Mudei para excluir tudos os lotes relacionado a documento nota fiscal/ctrc - furlan - 17/08/2015
    '            objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
    '            objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
    '            objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
    '            objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)

    '            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
    '                If chkIcms.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "ICMS" Then
    '                        enc.Base = BaseIcms.Text.Replace(".", "")
    '                        enc.Percentual = PerIcms.Text.Replace(".", "")
    '                        enc.Valor = ValorIcms.Text.Replace(".", "")
    '                    End If
    '                    enc.SituacaoTributaria = cstIcms.Text
    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   =  " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   =  " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         =  " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         =  " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    =  " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'ICMS'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorIcms.Text) AndAlso CDec(ValorIcms.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'ICMS', " & cstIcms.Text & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINS & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BaseIcms.Text) & "," & Str(PerIcms.Text) & "," & Str(PerIcms.Text) & "," & Str(ValorIcms.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If
    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributaria = " & cstIcms.Text & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkPis.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "PIS" Then
    '                        enc.Base = BasePis.Text.Replace(".", "")
    '                        enc.Percentual = PerPis.Text.Replace(".", "")
    '                        enc.Valor = ValorPis.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstPis.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'PIS'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorPis.Text) AndAlso CDec(ValorPis.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'PIS', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstPis.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BasePis.Text) & "," & Str(PerPis.Text) & "," & Str(PerPis.Text) & "," & Str(ValorPis.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstPis.Text & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkCofins.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "COFINS" Then
    '                        enc.Base = BaseCofins.Text
    '                        enc.Percentual = PerCofins.Text.Replace(".", "")
    '                        enc.Valor = ValorCofins.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstCofins.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINSOBS = txtObsCofins.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'COFINS'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorCofins.Text) AndAlso CDec(ValorCofins.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'COFINS', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstCofins.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BaseCofins.Text) & "," & Str(PerCofins.Text) & "," & Str(PerCofins.Text) & "," & Str(ValorCofins.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstCofins.Text & vbCrLf

    '                    If txtObsCofins.Text.Length > 0 Then
    '                        Sql &= ", ObservacaoPISCOFINS = " & txtObsCofins.Text & vbCrLf

    '                    End If

    '                    Sql &= "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkPCCR.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "PIS/COFINS/CSLL" Then
    '                        enc.Base = BasePCCR.Text.Replace(".", "")
    '                        enc.Percentual = PerPCCR.Text.Replace(".", "")
    '                        enc.Valor = ValorPCCR.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstPCCR.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'PIS/COFINS/CSLL'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorPCCR.Text) AndAlso CDec(ValorPCCR.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'PIS/COFINS/CSLL', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstPCCR.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BasePCCR.Text) & "," & Str(PerPCCR.Text) & "," & Str(PerPCCR.Text) & "," & Str(ValorPCCR.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)


    '                        Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstPCCR.Text & vbCrLf & _
    '                              "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                              "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                              "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                              "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                              "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                              "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                              "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                              "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                              "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                        Sqls.Add(Sql)
    '                    End If
    '                End If

    '                If chkPisR.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "PIS_RF" Then
    '                        enc.Base = BasePisR.Text.Replace(".", "")
    '                        enc.Percentual = PerPisR.Text.Replace(".", "")
    '                        enc.Valor = ValorPisR.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstPisR.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'PIS_RF'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorPisR.Text) AndAlso CDec(ValorPisR.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'PIS_RF', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstPisR.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BasePisR.Text) & "," & Str(PerPisR.Text) & "," & Str(PerPisR.Text) & "," & Str(ValorPisR.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstPisR.Text & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkCofinsR.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "COFINS_RF" Then
    '                        enc.Base = BaseCofinsR.Text
    '                        enc.Percentual = PerCofinsR.Text.Replace(".", "")
    '                        enc.Valor = ValorCofinsR.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstCofinsR.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINSOBS = txtObsCofinsR.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'COFINS_RF'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorCofinsR.Text) AndAlso CDec(ValorCofinsR.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'COFINS_RF', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstCofinsR.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BaseCofinsR.Text) & "," & Str(PerCofinsR.Text) & "," & Str(PerCofinsR.Text) & "," & Str(ValorCofinsR.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstCofinsR.Text & vbCrLf

    '                    If txtObsCofinsR.Text.Length > 0 Then
    '                        Sql &= ", ObservacaoPISCOFINS = " & txtObsCofinsR.Text & vbCrLf

    '                    End If

    '                    Sql &= "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkCSLLR.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "CSLL_RF" Then
    '                        enc.Base = BaseCSLLR.Text.Replace(".", "")
    '                        enc.Percentual = PerCSLLR.Text.Replace(".", "")
    '                        enc.Valor = ValorCSLLR.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstCSLLR.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'CSLL_RF'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorPisR.Text) AndAlso CDec(ValorPisR.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'CSLL_RF', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstCSLLR.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BaseCSLLR.Text) & "," & Str(PerCSLLR.Text) & "," & Str(PerCSLLR.Text) & "," & Str(ValorCSLLR.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstCSLLR.Text & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkPisCREPRE.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "PIS CRED PRE" Then
    '                        enc.Base = BasePisCREPRE.Text.Replace(".", "")
    '                        enc.Percentual = PerPisCREPRE.Text.Replace(".", "")
    '                        enc.Valor = ValorPisCREPRE.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstPisCREPRE.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'PIS CRED PRE'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorPisCREPRE.Text) AndAlso CDec(ValorPisCREPRE.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'PIS CRED PRE', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstPisCREPRE.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BasePisCREPRE.Text) & "," & Str(PerPisCREPRE.Text) & "," & Str(PerPisCREPRE.Text) & "," & Str(ValorPisCREPRE.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstPisCREPRE.Text & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If

    '                If chkCofinsCREPRE.Checked Then
    '                    'Carrega o objeto com os valores alterados
    '                    If enc.Codigo = "COFINS CRED PRE" Then
    '                        enc.Base = BaseCofinsCREPRE.Text
    '                        enc.Percentual = PerCofinsCREPRE.Text.Replace(".", "")
    '                        enc.Valor = ValorCofinsCREPRE.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINS = cstCofinsCREPRE.Text.Replace(".", "")
    '                        enc.SituacaoTributariaPISCOFINSOBS = txtObsCofinsCREPRE.Text.Replace(".", "")
    '                    End If

    '                    Sql = " Delete NotasFiscaisXEncargos " & vbCrLf & _
    '                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & vbCrLf & _
    '                          "    and Sequencia_Id    = " & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & vbCrLf & _
    '                          "    and Encargo_Id      = 'COFINS CRED PRE'"
    '                    Sqls.Add(Sql)

    '                    If Not String.IsNullOrWhiteSpace(ValorCofinsCREPRE.Text) AndAlso CDec(ValorCofinsCREPRE.Text) > 0 Then
    '                        Sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
    '                              "Cliente_Id, EndCliente_Id, " & vbCrLf & _
    '                              "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
    '                              "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf & _
    '                              "Encargo_Id, SituacaoTributaria, SituacaoTributariaIPI, SituacaoTributariaPISCOFINS, ObservacaoPISCOFINS, " & vbCrLf & _
    '                              "Operacao, SubOperacao, " & vbCrLf & _
    '                              "Grupo, Produto, " & vbCrLf & _
    '                              "EstadoOrigem, EstadoDestino, " & vbCrLf & _
    '                              "Base, Percentual, PercentualExibicao, Valor, CentroDeCusto)"
    '                        Sql &= "                          VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.CodigoCliente & "', " & objNotaFiscal.EnderecoCliente & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & objNotaFiscal.Serie & "', " & objNotaFiscal.Codigo & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "', " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Sequencia & " ," & vbCrLf & _
    '                               "'COFINS CRED PRE', " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributaria & ", " & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaIPI & "," & cstCofinsCREPRE.Text & "," & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).SituacaoTributariaPISCOFINSOBS & "," & vbCrLf & _
    '                               "" & objNotaFiscal.CodigoOperacao & ", " & objNotaFiscal.CodigoSubOperacao & ", " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoGrupoProduto & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CodigoProduto & "', " & vbCrLf & _
    '                               "'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoOrigem & "', '" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).EstadoDestino & "', " & vbCrLf & _
    '                               "" & Str(BaseCofinsCREPRE.Text) & "," & Str(PerCofinsCREPRE.Text) & "," & Str(PerCofinsCREPRE.Text) & "," & Str(ValorCofinsCREPRE.Text) & ",'" & objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos(0).CentroDeCusto & "')"
    '                        Sqls.Add(Sql)
    '                    End If

    '                    Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaPISCOFINS = " & cstCofinsCREPRE.Text & vbCrLf

    '                    If txtObsCofinsCREPRE.Text.Length > 0 Then
    '                        Sql &= ", ObservacaoPISCOFINS = " & txtObsCofinsCREPRE.Text & vbCrLf

    '                    End If

    '                    Sql &= "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                          "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                          "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                          "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                          "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                          "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                    Sqls.Add(Sql)
    '                End If
    '            Next

    '            If chkIPI.Checked Then
    '                Sql = " Update NotasFiscaisXEncargos set SituacaoTributariaIPI = " & cstIPI.Text & vbCrLf & _
    '                      "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                      "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                      "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                      "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                      "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                      "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                      "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
    '                      "    and Produto_Id      = '" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoProduto & "'" & vbCrLf & _
    '                      "    and CFOP_Id         = " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP
    '                Sqls.Add(Sql)
    '            End If


    '            Sql = " Update NotasFiscais Set ObservacoesControleInterno = '" & IIf(objNotaFiscal.ObservacoesControleInterno.Length > 0, objNotaFiscal.ObservacoesControleInterno & ". ACERTOS VIA ALTERAR ENCARGO - " & HttpContext.Current.Session("ssNomeUsuario") & "-" & Now().ToString("yyyy-MM-dd"), "ACERTOS VIA ALTERAR ENCARGO - " & HttpContext.Current.Session("ssNomeUsuario") & "-" & Now().ToString("yyyy-MM-dd")) & "'," & vbCrLf & _
    '                  "            UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
    '                  "        UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf & _
    '                  "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
    '                  "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
    '                  "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
    '                  "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
    '                  "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
    '                  "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
    '                  "    and Serie_Id        = '" & objNotaFiscal.Serie & "'"
    '            Sqls.Add(Sql)

    '            If Banco.GravaBanco(Sqls) Then
    '                objNotaFiscalCont = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

    '                Sqls.Clear()
    '                Dim alerta As String = ""

    '                'Temporário, apenas para Insol fazer os acertos necessários abaixo do ano de 2013.
    '                'If (HttpContext.Current.Session("ssNomeUsuario") = "YURI" Or HttpContext.Current.Session("ssNomeUsuario") = "CONTABILINSOL02") AndAlso objNotaFiscalCont.Movimento.Year <= 2014 Then
    '                '    alerta = " Yuri/CONTABILINSOL02, a Nota Fiscal não foi recontabilizada."
    '                'Else
    '                If objNotaFiscalCont.NFG Then
    '                    Dim cLote As Integer = 0
    '                    If objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.RPA) OrElse _
    '                        objNotaFiscalCont.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
    '                        cLote = 21
    '                    Else
    '                        cLote = 9
    '                    End If
    '                    objNotaFiscalCont.Razao.ContabilizarNotaSql(Sqls, cLote)
    '                Else
    '                    If objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CTRC) _
    '                            AndAlso objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ReciboDeFrete) _
    '                            AndAlso objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) _
    '                            AndAlso objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CT_E) _
    '                            AndAlso objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ComplementoDeFrete) _
    '                            AndAlso objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.RPA) _
    '                            AndAlso objNotaFiscalCont.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Anulacao) Then
    '                        objNotaFiscalCont.Razao.ContabilizarNotaSql(Sqls)
    '                    Else
    '                        objNotaFiscalCont.Razao.ContabilizarNotaSql(Sqls, 21)
    '                    End If
    '                End If
    '                'End If

    '                If Sqls.Count > 0 Then
    '                    If Banco.GravaBanco(Sqls) Then
    '                        objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscalCont)

    '                        SalvaNotaFiscal()

    '                        CarregarItem()
    '                        gridEncargos.Parent.Visible = False
    '                        gridEncargos.DataSource = Nothing
    '                        gridEncargos.DataBind()

    '                        BaseIcms.Text = ""
    '                        PerIcms.Text = ""
    '                        ValorIcms.Text = ""
    '                        BasePis.Text = ""
    '                        PerPis.Text = ""
    '                        ValorPis.Text = ""
    '                        BaseCofins.Text = ""
    '                        PerCofins.Text = ""
    '                        ValorCofins.Text = ""
    '                        BasePCCR.Text = ""
    '                        PerPCCR.Text = ""
    '                        ValorPCCR.Text = ""
    '                        BasePisR.Text = ""
    '                        PerPisR.Text = ""
    '                        ValorPisR.Text = ""
    '                        BaseCofinsR.Text = ""
    '                        PerCofinsR.Text = ""
    '                        ValorCofinsR.Text = ""
    '                        BaseCSLLR.Text = ""
    '                        PerCSLLR.Text = ""
    '                        ValorCSLLR.Text = ""
    '                        BaseCofinsCREPRE.Text = ""
    '                        PerCofinsCREPRE.Text = ""
    '                        ValorCofinsCREPRE.Text = ""
    '                        BasePisCREPRE.Text = ""
    '                        PerPisCREPRE.Text = ""
    '                        ValorPisCREPRE.Text = ""

    '                        cstIcms.Text = ""
    '                        cstPis.Text = ""
    '                        cstCofins.Text = ""
    '                        cstPCCR.Text = ""
    '                        cstPisR.Text = ""
    '                        cstCofinsR.Text = ""
    '                        cstCSLLR.Text = ""
    '                        cstCofinsCREPRE.Text = ""
    '                        cstPisCREPRE.Text = ""

    '                        txtObsCofins.Text = ""
    '                        txtObsCofinsR.Text = ""
    '                        txtObsCofinsCREPRE.Text = ""

    '                        chkIcms.Checked = False
    '                        chkIcms.Enabled = True
    '                        chkPis.Checked = False
    '                        chkCofins.Checked = False
    '                        chkIPI.Checked = False
    '                        chkPCCR.Checked = False
    '                        chkPisR.Checked = False
    '                        chkCofinsR.Checked = False
    '                        chkCSLLR.Checked = False
    '                        chkPisCREPRE.Checked = False
    '                        chkCofinsCREPRE.Checked = False

    '                        lnkAtualizar.Parent.Visible = False

    '                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
    '                    Else
    '                        MsgBox(Me.Page, "Erro ao Contabilizar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
    '                    End If
    '                Else
    '                    objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

    '                    SalvaNotaFiscal()

    '                    CarregarItem()

    '                    gridEncargos.Parent.Visible = False
    '                    gridEncargos.DataSource = Nothing
    '                    gridEncargos.DataBind()

    '                    BaseIcms.Text = ""
    '                    PerIcms.Text = ""
    '                    ValorIcms.Text = ""
    '                    BasePis.Text = ""
    '                    PerPis.Text = ""
    '                    ValorPis.Text = ""
    '                    BaseCofins.Text = ""
    '                    PerCofins.Text = ""
    '                    ValorCofins.Text = ""
    '                    BasePCCR.Text = ""
    '                    PerPCCR.Text = ""
    '                    ValorPCCR.Text = ""
    '                    BasePisR.Text = ""
    '                    PerPisR.Text = ""
    '                    ValorPisR.Text = ""
    '                    BaseCofinsR.Text = ""
    '                    PerCofinsR.Text = ""
    '                    ValorCofinsR.Text = ""
    '                    BaseCSLLR.Text = ""
    '                    PerCSLLR.Text = ""
    '                    ValorCSLLR.Text = ""
    '                    BaseCofinsCREPRE.Text = ""
    '                    PerCofinsCREPRE.Text = ""
    '                    ValorCofinsCREPRE.Text = ""
    '                    BasePisCREPRE.Text = ""
    '                    PerPisCREPRE.Text = ""
    '                    ValorPisCREPRE.Text = ""

    '                    cstIcms.Text = ""
    '                    cstPis.Text = ""
    '                    cstCofins.Text = ""
    '                    cstPCCR.Text = ""
    '                    cstPisR.Text = ""
    '                    cstCofinsR.Text = ""
    '                    cstCSLLR.Text = ""
    '                    cstCofinsCREPRE.Text = ""
    '                    cstPisCREPRE.Text = ""

    '                    txtObsCofins.Text = ""
    '                    txtObsCofinsR.Text = ""
    '                    txtObsCofinsCREPRE.Text = ""

    '                    chkIcms.Checked = False
    '                    chkIcms.Enabled = True
    '                    chkPis.Checked = False
    '                    chkCofins.Checked = False
    '                    chkIPI.Checked = False
    '                    chkPCCR.Checked = False
    '                    chkPisR.Checked = False
    '                    chkCofinsR.Checked = False
    '                    chkCSLLR.Checked = False

    '                    lnkAtualizar.Parent.Visible = False

    '                    MsgBox(Me.Page, "Registro alterado com Sucesso." & alerta, eTitulo.Sucess)
    '                End If
    '            Else
    '                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Private Sub IniciarAlteracaoNotaFiscal()

        RecuperaNotaFiscal()

        If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.DataNota, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento de Notas Fiscais já Fechado para esta data...")
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.DataNota, "CONTABIL") Then
            MsgBox(Me.Page, "Movimento Contábil já Fechado para esta data...")
        Else

            Dim original As [Lib].Negocio.NotaFiscal = objNotaFiscal.NotaFiscalOriginal

            If Not objNotaFiscal.SubOperacao.Devolucao AndAlso
                     objNotaFiscal.SubOperacao.Financeiro Then

                If objNotaFiscal.VencimentosNota Is Nothing Then
                    MsgBox(Me.Page, "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!", eTitulo.Info)
                    Exit Sub
                ElseIf objNotaFiscal.VencimentosNota.Count > 0 AndAlso Not objNotaFiscal.VencimentosNota.Sum(Function(s) s.ValorDoDocumento) = objNotaFiscal.TotalNota Then
                    MsgBox(Me.Page, "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!", eTitulo.Info)
                    Exit Sub
                End If
            End If

            If String.IsNullOrWhiteSpace(objNotaFiscal.ObservacoesControleInterno) Then
                objNotaFiscal.ObservacoesControleInterno = "ACERTOS VIA ALTERAR ENCARGO - " & HttpContext.Current.Session("ssNomeUsuario") & "-" & Now().ToString("yyyy-MM-dd")
            Else
                objNotaFiscal.ObservacoesControleInterno = objNotaFiscal.ObservacoesControleInterno & ". ACERTOS VIA ALTERAR ENCARGO - " & HttpContext.Current.Session("ssNomeUsuario") & "-" & Now().ToString("yyyy-MM-dd")
            End If

            If objNotaFiscal.NossaEmissao Then
                If objNotaFiscal.Salvar Then
                    Limpar()
                Else
                    MsgBox(Me.Page, "Erro ao Alterar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage" & HID.Value).ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            Else

                Dim Sqls As New ArrayList

                original.IUD = "D"
                original.CarregandoNota = True
                original.SalvarSql(Sqls)

                objNotaFiscal.IUD = "I"

                objNotaFiscal.CarregandoNota = True
                If String.IsNullOrWhiteSpace(txtSerie.Text) Then
                    MsgBox(Me.Page, "Série Nota Fiscal não foi informada.")
                    Exit Sub
                End If
                objNotaFiscal.Serie = Trim(txtSerie.Text)
                If txtSerie.Text.Length > 0 Then objNotaFiscal.SerieNotaProdutor = Trim(txtSerie.Text)
                txtSerie.Text = objNotaFiscal.Serie
                objNotaFiscal.UsuarioInclusao = original.UsuarioInclusao
                objNotaFiscal.DataInclusao = original.DataInclusao
                objNotaFiscal.UsuarioAlteracao = Session("ssNomeUsuario")
                objNotaFiscal.DataAlteracao = Date.Now

                If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                    For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                        dr.IUD = "I"
                    Next
                    objNotaFiscal.VencimentosNota.NF = objNotaFiscal
                    'Atualiza a provisão do pedido.
                    For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosPedido.Where(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                        dr.IUD = IIf(dr.IUD = "I", "I", "U")
                        If dr.IUD = "U" Then dr.SalvarSql(Sqls, False)
                    Next

                End If

                objNotaFiscal.SalvarSql(Sqls)

                If Banco.GravaBanco(Sqls) Then
                    Limpar()
                Else
                    MsgBox(Me.Page, "Erro ao Alterar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage" & HID.Value).ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        End If

    End Sub

    Protected Sub gridItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperaNotaFiscal()

            'gridEncargos.Parent.Visible = True
            'gridEncargos.DataSource = objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.ToArray
            'gridEncargos.DataBind()

            lnkAtualizar.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            objNotaFiscal.CodigoEmpresa = Empresa(0)
            objNotaFiscal.DataNota = txtDataInicial.Text
            objNotaFiscal.Movimento = txtDataFinal.Text
            If Not String.IsNullOrWhiteSpace(txtES.Text) Then objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", 0, 1)
            If Not String.IsNullOrWhiteSpace(txtSerie.Text) Then objNotaFiscal.Serie = txtSerie.Text
            If Not String.IsNullOrWhiteSpace(txtNota.Text) Then objNotaFiscal.Codigo = txtNota.Text
            objNotaFiscal.CodigoSituacao = 1
            objNotaFiscal.CodigoTipoDeDocumento = 0

            Session("objNotaFiscal" & HID.Value) = objNotaFiscal
            Session("ssCampo" & HID.Value) = "AlterarEncargosCTE"
            ucConsultaPedidosXNotas.SetarHID(HID.Value)
            ucConsultaPedidosXNotas.BindGridView()
            Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNXI")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarEncargo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class