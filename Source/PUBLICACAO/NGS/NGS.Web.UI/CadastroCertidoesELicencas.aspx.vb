Imports NGS.Lib.Negocio

Public Class CadastroCertidoesELicencas
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CadastroCertidoesELicencas", "ACESSAR") Then
                    LimparCampos()
                    ddl.Carregar(ddlTipo, CarregarDDL.Tabela.TipoDeCertidao)
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas)
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    HID.Value = Guid.NewGuid().ToString
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LimparCampos()
        ddlTipo.Text = String.Empty
        txtDataEm.Text = String.Empty
        txtDataVenc.Text = String.Empty
        txtEmail.Text = String.Empty
        txtAviso.Text = String.Empty
        txtObservacao.Text = String.Empty

        GridCadastroCertidaoELicenca.DataBind()

        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        ddlEmpresa.Enabled = True
        ddlTipo.Enabled = True
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlTipo.SelectedValue) Then
            MsgBox(Me.Page, "Tipo de certidão ou licença não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa não foi informada!", eTitulo.Info)
            Return False
        End If

        Return True
    End Function

    Private Function verificaDatas() As Boolean
        If String.IsNullOrWhiteSpace(txtDataEm.Text) OrElse String.IsNullOrWhiteSpace(txtDataVenc.Text) _
           OrElse Not IsDate(txtDataEm.Text) OrElse Not IsDate(txtDataVenc.Text) Then
            MsgBox(Me.Page, "Informe a data de emissão e a data de vencimento.")
            Return False
        ElseIf CDate(txtDataEm.Text) > CDate(txtDataVenc.Text) Then
            MsgBox(Me.Page, "Data de emissão não pode ser maior que a data de vencimento.")
            Return False
        End If
        Return True
    End Function

    Private Function PreenhcerObjetos(ByVal IUD As String) As CadastroCertidaoELicenca
        Dim obj As New [Lib].Negocio.CadastroCertidaoELicenca()

        With obj
            .IUD = IUD
            .CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0).Trim
            .DescEmp = ddlEmpresa.SelectedValue.Split("-")(1).Trim
            .CodigoTipo = CInt(ddlTipo.SelectedValue.Trim)
            .DataEmissao = CDate(txtDataEm.Text.Trim)
            .DataVencimento = CDate(txtDataVenc.Text.Trim)
            .Email = txtEmail.Text.Trim
            .AvisoDias = CInt(txtAviso.Text.Trim)
            .Observacao = txtObservacao.Text.Trim
        End With

        Return obj
    End Function

    Public Overrides Sub Carregar(ByVal str As String)
        ddl.Carregar(ddlTipo, CarregarDDL.Tabela.TipoDeCertidao)
    End Sub

    Protected Sub GridCadastroCertidaoELicenca_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridCadastroCertidaoELicenca.SelectedIndexChanged
        Try
            ddlTipo.SelectedValue = Server.HtmlDecode(GridCadastroCertidaoELicenca.SelectedRow.Cells(1).Text.Split("-")(0).Trim)
            txtDataEm.Text = Server.HtmlDecode(GridCadastroCertidaoELicenca.SelectedRow.Cells(2).Text())
            txtDataVenc.Text = Server.HtmlDecode(GridCadastroCertidaoELicenca.SelectedRow.Cells(3).Text())
            txtEmail.Text = Server.HtmlDecode(GridCadastroCertidaoELicenca.SelectedRow.Cells(4).Text())
            txtAviso.Text = Server.HtmlDecode(GridCadastroCertidaoELicenca.SelectedRow.Cells(5).Text())
            txtObservacao.Text = Server.HtmlDecode(GridCadastroCertidaoELicenca.SelectedRow.Cells(6).Text())

            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            lnkNovo.Parent.Visible = False

            ddlEmpresa.Enabled = False
            ddlTipo.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CadastroCertidoesELicencas", "GRAVAR") Then
                If ValidarCampos() Then
                    If verificaDatas() Then
                        Dim obj As [Lib].Negocio.CadastroCertidaoELicenca = PreenhcerObjetos("I")

                        If obj.Salvar() Then
                            MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
                            LimparCampos()
                        Else
                            MsgBox(Me.Page, "Erro durante o processo.")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CadastroCertidoesELicencas", "ALTERAR") Then
                If ValidarCampos() Then
                    If verificaDatas() Then
                        Dim obj As [Lib].Negocio.CadastroCertidaoELicenca = PreenhcerObjetos("U")

                        If obj.Salvar() Then
                            MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                            LimparCampos()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CadastroCertidoesELicencas", "EXCLUIR") Then
                If ValidarCampos() Then
                    If verificaDatas() Then
                        Dim obj As New [Lib].Negocio.CadastroCertidaoELicenca
                        obj.IUD = "D"
                        obj.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)


                        If obj.Salvar() Then
                            MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
                            LimparCampos()
                        Else
                            MsgBox(Me.Page, "Erro durante o processo.")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("CadastroCertidoesELicencas", "LEITURA") Then
                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    Dim lst As New [Lib].Negocio.ListCadastroCertidaoELicenca(ddlEmpresa.SelectedValue.Split("-")(0), ddlTipo.SelectedValue)

                    GridCadastroCertidaoELicenca.DataSource = lst
                    GridCadastroCertidaoELicenca.DataBind()
                Else
                    MsgBox(Me.Page, "Empresa é obrigatória.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CadastroCertidoesELicencas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkAddTipo_Click(sender As Object, e As EventArgs) Handles lnkAddTipo.Click
        Try
            ucCadastrarTipoDeCertidao.Limpar()
            ucCadastrarTipoDeCertidao.SetarHID(HID.Value)
            Popup.CadastrarTipoDeCertidao(Me.Page, "CadastrarTipoDeCertidao" & HID.Value.ToString(), "txtDescricao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class