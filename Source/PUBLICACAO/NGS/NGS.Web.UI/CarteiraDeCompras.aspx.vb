Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CarteiraDeCompras
    Inherits BasePage

#Region "Atributos / Propriedades"

    Dim consultaTabelas As New ConsultaTabelas
    Dim linha As String = ""
    Dim cl As Integer = 0
    Dim pagina As Integer = 0
    Dim Sql As String = ""

#End Region

#Region "Eventos"

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CarteiraDeCompras", "ACESSAR") Then
                LimparPagina()
                Limpar()
                CarregarDropDownList()
                BindGridView()
                GerenciarBotoes()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub imbContaCliente_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbContaCliente.Click
        Try
            hdnContaSelecionada.Value = "CLIENTE"
            ucConsultaContaCliente.BindGridView("CLIENTE")
            Popup.ConsultaDeContaCliente(Me.Page, "objCliente" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbPesquisarDesconto_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbPesquisarDesconto.Click
        Try
            hdnContaSelecionada.Value = "DESCONTO"
            ucConsultaContaCliente.BindGridView("DESCONTO")
            Popup.ConsultaDeContaCliente(Me.Page, "objDesconto" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbPesquisarDeducoes_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbPesquisarDeducoes.Click
        Try
            hdnContaSelecionada.Value = "DEDUCOES"
            ucConsultaContaCliente.BindGridView("DEDUCOES")
            Popup.ConsultaDeContaCliente(Me.Page, "objDeducoes" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbPesquisarJuros_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbPesquisarJuros.Click
        Try
            hdnContaSelecionada.Value = "JUROS"
            ucConsultaContaCliente.BindGridView("JUROS")
            Popup.ConsultaDeContaCliente(Me.Page, "objJuros" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbPesquisarAcrescimos_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbPesquisarAcrescimos.Click
        Try
            hdnContaSelecionada.Value = "ACRESCIMOS"
            ucConsultaContaCliente.BindGridView("ACRESCIMOS")
            Popup.ConsultaDeContaCliente(Me.Page, "objAcrescimos" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbLimparContaCliente_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbLimparContaCliente.Click
        Try
            txtContaCliente.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbLimparDesconto_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbLimparDesconto.Click
        Try
            txtContaDesconto.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbLimparDeducooes_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbLimparDeducooes.Click
        Try
            txtContaDeducoes.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbLimparJuros_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbLimparJuros.Click
        Try
            txtContaJuros.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imbLimparAcrescimos_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imbLimparAcrescimos.Click
        Try
            txtContaAcrescimos.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gdvCarteiraDeCompras_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gdvCarteiraDeCompras.SelectedIndexChanged
        Try
            txtCodigo.Text = gdvCarteiraDeCompras.SelectedRow.Cells(1).Text
            txtCodigo.ReadOnly = True
            txtCodigo.Enabled = False
            txtDescricao.Text = Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(2).Text)
            txtContaCliente.Text = Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(4).Text).Trim & " - " & Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(10).Text).Trim
            txtContaDesconto.Text = Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(5).Text).Trim & " - " & Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(11).Text).Trim
            txtContaDeducoes.Text = Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(6).Text).Trim & " - " & Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(12).Text).Trim
            txtContaJuros.Text = Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(7).Text).Trim & " - " & Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(13).Text).Trim
            txtContaAcrescimos.Text = Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(8).Text).Trim & " - " & Server.HtmlDecode(gdvCarteiraDeCompras.SelectedRow.Cells(14).Text).Trim
            ddlSituacao.SelectedIndex = gdvCarteiraDeCompras.SelectedRow.Cells(3).Text
            ddlClassificacao.SelectedValue = gdvCarteiraDeCompras.SelectedRow.Cells(9).Text
            If (gdvCarteiraDeCompras.SelectedRow.Cells(15).Text.Equals("S")) Then
                radSim.Checked = True
                radNao.Checked = False
                radSim_CheckedChanged(radSim, New EventArgs)
            Else
                radNao.Checked = True
                radSim.Checked = False
                radNao_CheckedChanged(radNao, New EventArgs)
            End If
            chkPedido.Checked = IIf(gdvCarteiraDeCompras.SelectedRow.Cells(16).Text().Equals("S"), True, False)
            chkBaixaDeAdiantamento.Checked = IIf(gdvCarteiraDeCompras.SelectedRow.Cells(17).Text().Equals("S"), True, False)
            GerenciarBotoes()
            ValidarConta()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ValidarConta()
        Try
            Dim sql As String = " select top 1 Carteira from ContasAPagar " & vbCrLf &
                                " where Carteira = '" & txtCodigo.Text & "'" & vbCrLf &
                                " union " & vbCrLf &
                                " select top 1 Carteira from ContasAReceber " & vbCrLf &
                                " where Carteira = '" & txtCodigo.Text & "'" & vbCrLf &
                                " union " & vbCrLf &
                                " select top 1 Carteira from MovimentacoesFinanceiras " & vbCrLf &
                                " where Carteira = '" & txtCodigo.Text & "'"
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Carteira")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False

                MsgBox(Me.Page, "Finalidade Financeira não pode ser alterada/excluída pois está sendo usada no Financeiro.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDeCompras", "GRAVAR") Then
                Dim Erro As String = ValidarCampos()
                If (String.IsNullOrEmpty(Erro)) Then
                    Dim adiantamento As String
                    If (radSim.Checked) Then
                        adiantamento = "S"
                    Else
                        adiantamento = "N"
                    End If
                    Incluir(txtCodigo.Text,
                            txtDescricao.Text,
                            ddlSituacao.SelectedValue,
                            ddlClassificacao.SelectedValue,
                            txtContaCliente.Text.Split("-").ToArray(0).Trim,
                            txtContaDesconto.Text.Split("-").ToArray(0).Trim,
                            txtContaDeducoes.Text.Split("-").ToArray(0).Trim,
                            txtContaJuros.Text.Split("-").ToArray(0).Trim,
                            txtContaAcrescimos.Text.Split("-").ToArray(0).Trim,
                            adiantamento,
                            chkPedido.Checked,
                            chkBaixaDeAdiantamento.Checked)
                    MsgBox(Me.Page, "Cadastro realizado com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, Erro)
                End If
            Else
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDeCompras", "ALTERAR") Then
                Dim Erro As String = ValidarCampos()
                If (String.IsNullOrEmpty(Erro)) Then
                    Dim adiantamento As String
                    If (radSim.Checked) Then
                        adiantamento = "S"
                    Else
                        adiantamento = "N"
                    End If
                    Alterar(txtCodigo.Text,
                            txtDescricao.Text,
                            ddlSituacao.SelectedValue,
                            ddlClassificacao.SelectedValue,
                            txtContaCliente.Text.Split("-").ToArray(0).Trim,
                            txtContaDesconto.Text.Split("-").ToArray(0).Trim,
                            txtContaDeducoes.Text.Split("-").ToArray(0).Trim,
                            txtContaJuros.Text.Split("-").ToArray(0).Trim,
                            txtContaAcrescimos.Text.Split("-").ToArray(0).Trim,
                            adiantamento,
                            chkPedido.Checked,
                            chkBaixaDeAdiantamento.Checked)
                    MsgBox(Me.Page, "Alteração realizada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, Erro)
                End If
            Else
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDeCompras", "EXCLUIR") Then
                Dim Erro As String = ValidarCampos()
                If (String.IsNullOrEmpty(Erro)) Then
                    Excluir(txtCodigo.Text)
                    MsgBox(Me.Page, "Exclusão realizada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, Erro)
                End If
            Else
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            Relatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Sub LimparPagina()
        Session.Remove("objPlanoDeContas" & HID.Value.ToString)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaPlanoDeContas.SetarHID(HID.Value)
    End Sub

    Private Sub CarregarDropDownList()
        ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "", True)
        ddlSituacao.SelectedIndex = "0"
        ddlClassificacao.Items.Clear()
        Dim Descricao As String
        Descricao = "P - Contas a Pagar"
        ddlClassificacao.Items.Add(New ListItem(Descricao, "P"))
        Descricao = "R - Contas a Receber"
        ddlClassificacao.Items.Add(New ListItem(Descricao, "R"))
        Descricao = "M - Movimentações Financeiras"
        ddlClassificacao.Items.Add(New ListItem(Descricao, "M"))
    End Sub

    Public Overrides Sub Carregar(str As String)
        Select Case hdnContaSelecionada.Value
            Case "CLIENTE"
                txtContaCliente.Text = Server.HtmlDecode(str)
            Case "DESCONTO"
                txtContaDesconto.Text = Server.HtmlDecode(str)
            Case "DEDUCOES"
                txtContaDeducoes.Text = Server.HtmlDecode(str)
            Case "JUROS"
                txtContaJuros.Text = Server.HtmlDecode(str)
            Case "ACRESCIMOS"
                txtContaAcrescimos.Text = Server.HtmlDecode(str)
        End Select
    End Sub

    Private Sub BindGridView()
        Dim ds As New DataSet
        Sql = "SELECT CxP.Produto_Id, CxP.Descricao, CxP.Situacao, " & vbCrLf & _
              "       ISNULL(CxP.ContaCLientes, N'') AS ContaClientes, " & vbCrLf & _
              "       ISNULL(CxP.ContaDescontos, N'') AS ContaDescontos, ISNULL(CxP.ContaDeducoes, N'') AS ContaDeducoes, " & vbCrLf & _
              "       ISNULL(CxP.ContaJuros, N'') AS ContaJuros, ISNULL(CxP.ContaAcrescimos, N'') AS ContaAcrescimos, " & vbCrLf & _
              "       ISNULL(CxP.Classificacao, N'') AS Classificacao, ISNULL(PlanoClientes.Titulo, N'') AS NomeContaCliente, ISNULL(PlanoDescontos.Titulo, N'') " & vbCrLf & _
              "       AS NomeContaDesconto, ISNULL(PlanoDeducoes.Titulo, N'') AS NomeContaDeducao, ISNULL(PlanoJuros.Titulo, N'') AS NomeContaJuro, " & vbCrLf & _
              "       ISNULL(PlanoAcrescimos.Titulo, N'') AS NomeContaAcrescimo, ISNULL(CxP.Adiantamento, 'N') AS Adiantamento, " & vbCrLf & _
              "       CASE WHEN CxP.Pedido = 1 THEN 'S' ELSE 'N' END Pedido , " & vbCrLf & _
              "       CASE WHEN CxP.BaixaAdiantamento = 1 THEN 'S' ELSE 'N' END BaixaAdiantamento, " & vbCrLf & _
              "       ISNULL(CxP.CarteiraBaixaAdiantamento, N'') CarteiraBaixaAdiantamento  " & vbCrLf & _
              "   FROM ComprasXProdutos CxP " & vbCrLf & _
              "   LEFT JOIN PlanoDeContas AS PlanoAcrescimos " & vbCrLf & _
              "     ON CxP.ContaAcrescimos = PlanoAcrescimos.Conta_Id " & vbCrLf & _
              "   LEFT JOIN PlanoDeContas AS PlanoJuros " & vbCrLf & _
              "     ON CxP.ContaJuros = PlanoJuros.Conta_Id  " & vbCrLf & _
              "   LEFT JOIN PlanoDeContas AS PlanoDeducoes " & vbCrLf & _
              "     ON CxP.ContaDeducoes = PlanoDeducoes.Conta_Id " & vbCrLf & _
              "   LEFT JOIN PlanoDeContas AS PlanoDescontos " & vbCrLf & _
              "     ON CxP.ContaDescontos = PlanoDescontos.Conta_Id " & vbCrLf & _
              "   LEFT JOIN PlanoDeContas AS PlanoClientes " & vbCrLf & _
              "     ON CxP.ContaCLientes = PlanoClientes.Conta_Id " & vbCrLf & _
              "  ORDER BY CxP.Produto_Id" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "ComprasXProdutos")
        gdvCarteiraDeCompras.DataSource = ds
        gdvCarteiraDeCompras.DataBind()
    End Sub

    Private Sub Relatorio()
        Try
            If Funcoes.VerificaPermissao("CarteiraDeCompras", "RELATORIO") Then
                Dim Ds_CarteiraDeCompras As New DataSet

                Sql = "SELECT Produto_Id AS Codigo, Descricao, ISNULL(Situacao, '') AS Situacao, ISNULL(ContaCLientes, '') AS ContaCLientes, " & vbCrLf & _
                      "ISNULL(ContaDescontos, '') AS ContaDescontos, ISNULL(ContaDeducoes, '') AS ContaDeducoes, " & vbCrLf & _
                      "ISNULL(ContaJuros, '') AS ContaJuros, ISNULL(ContaAcrescimos, '') AS ContaAcrescimos, Classificacao " & vbCrLf & _
                      "FROM ComprasXProdutos " & vbCrLf & _
                      "ORDER BY Produto_Id" & vbCrLf
                Ds_CarteiraDeCompras = Banco.ConsultaDataSet(Sql, "CarteiraDeCompras")

                Funcoes.BindReport(Me.Page, Ds_CarteiraDeCompras, "Cr_CarteiraDeCompras", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            param &= "Código: " & txtCodigo.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) Then
            param &= "Situação: " & ddlSituacao.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlClassificacao.SelectedValue) Then
            param &= "Classificação: " & ddlClassificacao.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtContaCliente.Text) Then
            param &= "Conta Clientes: " & txtContaCliente.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtContaDesconto.Text) Then
            param &= "Conta Descontos: " & txtContaDesconto.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtContaDeducoes.Text) Then
            param &= "Conta Deduções: " & txtContaDeducoes.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtContaJuros.Text) Then
            param &= "Conta Juros: " & txtContaJuros.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtContaAcrescimos.Text) Then
            param &= "Conta Acréscismos: " & txtContaAcrescimos.Text & vbCrLf
        End If
        param &= IIf(radSim.Checked, "Adiantamento: Sim", "Adiantamento: Não")
        param &= "Pedido: " & (IIf(chkPedido.Checked, "Sim ", "Não "))
        param &= "Baixa de Adiantamento: " & (IIf(chkBaixaDeAdiantamento.Checked, "Sim ", "Não "))

        Return param
    End Function

    Private Function ValidarCampos()
        Dim Erro As String = String.Empty
        If (String.IsNullOrWhiteSpace(txtCodigo.Text)) Then
            Erro &= "O campo Código é obrigatório. \n"
        End If
        If (String.IsNullOrWhiteSpace(txtDescricao.Text)) Then
            Erro &= "O campo Descrição é obrigatório. \n"
        End If
        If (ddlSituacao.SelectedValue.Equals(String.Empty)) Then
            Erro &= "O campo Situação é obrigatório. \n"
        End If
        If (String.IsNullOrEmpty(ddlClassificacao.SelectedItem.Text)) Then
            Erro &= "O campo Classificação é obrigatório! \n"
        End If
        If (Not radSim.Checked And Not radNao.Checked) Then
            Erro &= "O campo Adiantamento é obrigatório. \n"
        End If
        Return Erro
    End Function

    Protected Sub Limpar()
        Try
            txtCodigo.Text = String.Empty
            txtDescricao.Text = String.Empty
            ddlSituacao.ClearSelection()
            ddlClassificacao.ClearSelection()
            txtContaCliente.Text = String.Empty
            txtContaDesconto.Text = String.Empty
            txtContaDeducoes.Text = String.Empty
            txtContaJuros.Text = String.Empty
            txtContaAcrescimos.Text = String.Empty
            txtCodigo.ReadOnly = False
            txtCodigo.Enabled = True
            radSim.Checked = False
            radNao.Checked = True
            radNao_CheckedChanged(radNao, New EventArgs)
            chkPedido.Checked = False
            chkBaixaDeAdiantamento.Checked = False
            GerenciarBotoes()
            BindGridView()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function Incluir(ByVal Codigo As String, ByVal Descricao As String, ByVal Situacao As String, ByVal Classificacao As String, _
                            ByVal Conta As String, ByVal Conta1 As String, ByVal Conta2 As String, ByVal Conta3 As String, ByVal Conta4 As String, _
                            ByVal Adiantamento As String, Pedido As Boolean, BaixaAdiantamento As Boolean)
        If Funcoes.VerificaPermissao("CarteiraDeCompras", "GRAVAR") Then
            Dim SqlArray As New ArrayList

            Sql = " INSERT INTO ComprasXProdutos (Produto_Id, Descricao, Situacao, ContaCLientes, ContaDescontos, " & vbCrLf & _
                  "        ContaDeducoes, ContaJuros, ContaAcrescimos, Classificacao, UsuarioInclusao, UsuarioInclusaoData, Adiantamento," & vbCrLf & _
                  "        Pedido, BaixaAdiantamento ) " & vbCrLf & _
                  " VALUES ('" & Codigo & "','" & Descricao & "','" & Situacao & "','" & Conta & "','" & Conta1 & "','" & Conta2 & "'" & vbCrLf & _
                  "         ,'" & Conta3 & "','" & Conta4 & "','" & Classificacao & "','" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                  "         ,CONVERT(DATETIME,'" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "',102), '" & Adiantamento & "'" & vbCrLf & _
                  "          ,'" & Pedido & "','" & BaixaAdiantamento & "' )" & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                Return HttpContext.Current.Session("ssMessage")
            Else
                Return ""
            End If

        Else
            Return "Usuário sem permissão para Gravar Registro"
        End If

    End Function

    Public Function Alterar(ByVal Codigo As String, ByVal Descricao As String, ByVal Situacao As String, ByVal Classificacao As String, _
                            ByVal Conta As String, ByVal Conta1 As String, ByVal Conta2 As String, ByVal Conta3 As String, ByVal Conta4 As String, _
                            ByVal Adiantamento As String, Pedido As Boolean, BaixaAdiantamento As Boolean)
        If Funcoes.VerificaPermissao("CarteiraDeCompras", "ALTERAR") Then
            Dim SqlArray As New ArrayList

            Sql = "Update ComprasXProdutos " & vbCrLf & _
                  "   Set Descricao = '" & Descricao & "', " & vbCrLf & _
                  "       Situacao = '" & Situacao & "', " & vbCrLf & _
                  "       ContaCLientes = '" & Conta & "', " & vbCrLf & _
                  "       ContaDescontos = '" & Conta1 & "', " & vbCrLf & _
                  "       ContaDeducoes = '" & Conta2 & "', " & vbCrLf & _
                  "       ContaJuros = '" & Conta3 & "', " & vbCrLf & _
                  "       ContaAcrescimos = '" & Conta4 & " ', " & vbCrLf & _
                  "       Classificacao = '" & Classificacao & "', " & vbCrLf & _
                  "       UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf & _
                  "       UsuarioAlteracaoData = CONVERT(DATETIME,'" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "',102), " & vbCrLf & _
                  "       Adiantamento = '" & Adiantamento & "', " & vbCrLf & _
                  "       Pedido = '" & Pedido & "'," & vbCrLf & _
                  "       BaixaAdiantamento  = '" & BaixaAdiantamento & "'" & vbCrLf & _
                  " WHERE Produto_Id = '" & Codigo & "'"
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                Return HttpContext.Current.Session("ssMessage")
            Else
                Return ""
            End If
        Else
            Return "Usuário sem permissão para Alterar Registro"
        End If

    End Function

    Public Function Excluir(ByVal Codigo As String)
        If Funcoes.VerificaPermissao("CarteiraDeCompras", "EXCLUIR") Then
            Dim SqlArray As New ArrayList

            Sql = "Delete From ComprasXProdutos Where Produto_Id = '" & Codigo & "'"
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                Return HttpContext.Current.Session("ssMessage")
            Else
                Return ""
            End If
        Else
            Return "Usuário sem permissão para Excluir Registro"
        End If
    End Function

    Public Sub GerenciarBotoes()
        If (String.IsNullOrWhiteSpace(txtCodigo.Text)) Then
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            lnkNovo.Parent.Visible = True
            ddlSituacao.SelectedIndex = 1
            ddlSituacao.Enabled = False
        Else
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            lnkNovo.Parent.Visible = False
            ddlSituacao.Enabled = True
        End If
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CarteiraDeCompras")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub radSim_CheckedChanged(sender As Object, e As EventArgs) Handles radSim.CheckedChanged
        divBaixaDeAdiantamento.Style.Clear()
    End Sub

    Protected Sub radNao_CheckedChanged(sender As Object, e As EventArgs) Handles radNao.CheckedChanged
        divBaixaDeAdiantamento.Style.Add("Display", "none")
    End Sub
End Class