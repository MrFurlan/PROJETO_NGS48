Imports NGS.Lib.Negocio

Public Class TipoDeCertidaoLicenca
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack AndAlso IsConnect Then
                If Funcoes.VerificaPermissao("CadastroTipoDeCertidaoELicenca", "ACESSAR") Then
                    CarregarGrid()
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarGrid()
        If Funcoes.VerificaPermissao("CadastroTipoDeCertidaoELicenca", "LEITURA") Then
            Dim lst As New [Lib].Negocio.ListTipoDeCertidao(True)

            GridCadastroTipoDeCertidaoELicenca.DataSource = lst
            GridCadastroTipoDeCertidaoELicenca.DataBind()
        End If
    End Sub

    Private Sub LimparCampos()
        txtCodigo.Text = String.Empty
        txtDescricao.Text = String.Empty

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe a descrição!", eTitulo.Info)
            Return False
        End If

        Return True
    End Function

    Private Function PreenhcerObjetos(ByVal IUD As String) As TipoDeCertidao
        Dim obj As New [Lib].Negocio.TipoDeCertidao

        With obj
            .IUD = IUD
            .Codigo = IIf(IUD = "I", 0, txtCodigo.Text)
            .Descricao = txtDescricao.Text
        End With

        Return obj
    End Function

    Protected Sub GridCadastroTipoDeCertidaoELicenca_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridCadastroTipoDeCertidaoELicenca.SelectedIndexChanged
        Try
            txtCodigo.Text = Server.HtmlDecode(GridCadastroTipoDeCertidaoELicenca.SelectedRow.Cells(1).Text)
            txtDescricao.Text = Server.HtmlDecode(GridCadastroTipoDeCertidaoELicenca.SelectedRow.Cells(2).Text)

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CadastroTipoDeCertidaoELicenca", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim obj As [Lib].Negocio.TipoDeCertidao = PreenhcerObjetos("I")

                    If obj.Salvar() Then
                        MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
                        LimparCampos()
                        CarregarGrid()
                    Else
                        MsgBox(Me.Page, "Erro durante o processo.")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CadastroTipoDeCertidaoELicenca", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim obj As [Lib].Negocio.TipoDeCertidao = PreenhcerObjetos("U")

                    If obj.Salvar() Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        LimparCampos()
                        CarregarGrid()
                    Else
                        MsgBox(Me.Page, "Erro durante o processo.")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CadastroTipoDeCertidaoELicenca", "EXCLUIR") Then
                If ValidarCampos() Then
                    Dim obj As New [Lib].Negocio.TipoDeCertidao
                    obj.IUD = "D"
                    obj.Codigo = txtCodigo.Text

                    If obj.Salvar() Then
                        MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
                        LimparCampos()
                        CarregarGrid()
                    Else
                        MsgBox(Me.Page, "Erro durante o processo.")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CadastroTipoDeCertidaoELicenca")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class