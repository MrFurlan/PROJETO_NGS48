Imports NGS.Lib.Negocio

Public Class AlterarObservacaoLaudo
    Inherits BasePage
    Dim objLaudo As Pesagem

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AlterarLaudo", "ACESSAR") Then
                    LimparCampos()
                    CarregaUnidadeDeNegocio()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SessaoRecuperaLaudo()
        objLaudo = CType(Session("objLaudo" & HID.Value), Pesagem)
        If objLaudo Is Nothing Then objLaudo = New Pesagem()
    End Sub
    Private Sub SessaoSalvaLaudo()
        Session("objLaudo" & HID.Value) = objLaudo
    End Sub

    Private Sub LimparCampos()
        txtLaudo.Text = String.Empty
        txtObservacao.Text = String.Empty
        txtCliente.Text = String.Empty

        Session.Remove("objLaudo" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        ddlUnidadeDeNegocio.Enabled = True
        ddlEmpresa.Enabled = True
        txtLaudo.Enabled = True

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtLaudo.Text) Then
            MsgBox(Me.Page, "Campo laudo obrigatorio.")
            Return False
        End If

        Return True
    End Function

    Private Sub CarregaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CarregaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            CarregaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarObservacaoLaudo", "ALTERAR") Then
                SessaoRecuperaLaudo()

                Dim sqls As New ArrayList

                objLaudo.IUD = "U"
                objLaudo.Observacoes = txtObservacao.Text

                Dim sql As String = "Update Pesagem set " & vbCrLf & _
                                    "       UsuarioAlteracao        ='" & Session("ssNomeUsuario").ToString & "'" & vbCrLf & _
                                    "       ,UsuarioAlteracaoData    = '" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                                    "       ,Observacoes             ='" & objLaudo.Observacoes & "'" & vbCrLf & _
                                    " Where Empresa_Id    ='" & objLaudo.CodigoEmpresa & "'" & vbCrLf & _
                                    "   and EndEmpresa_Id = " & objLaudo.EnderecoEmpresa & vbCrLf & _
                                    "   and Pesagem_Id    = " & objLaudo.Codigo & vbCrLf & _
                                    "   and Sequencia_Id  = " & objLaudo.Sequencia

                sqls.Add(sql)

                If Banco.GravaBanco(sqls) Then
                    MsgBox(Me.Page, "Observação alterada com sucesso.")
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Erro a Altererar o Laudo de Pesagem.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarObservacaoLaudo", "CONSULTAR") Then
                If ValidarCampos() Then
                    SessaoRecuperaLaudo()

                    Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

                    objLaudo = New Pesagem(Empresa(0), Empresa(1), txtLaudo.Text)

                    txtCliente.Text = objLaudo.Cliente.Codigo & "-" & objLaudo.Cliente.Nome

                    If String.IsNullOrWhiteSpace(objLaudo.Observacoes) Then
                        txtObservacao.Text = ""
                    Else
                        txtObservacao.Text = objLaudo.Observacoes
                    End If

                    lnkAtualizar.Parent.Visible = True
                    lnkConsultar.Parent.Visible = False

                    ddlUnidadeDeNegocio.Enabled = False
                    ddlEmpresa.Enabled = False
                    txtLaudo.Enabled = False

                    SessaoSalvaLaudo()
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

        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class