Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucSupervisor
    Inherits BaseUserControl

    Public Property Processo() As String
        Get
            Return Session("Processo")
        End Get
        Set(ByVal value As String)
            Session("Processo") = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub btnEntrar_Click(sender As Object, e As EventArgs) Handles btnEntrar.Click
        Dim ObjUsuario = New Usuario(txtUsuario.Text)
        If Not UsuarioServidor.Validar(txtUsuario.Text, txtSenha.Text, True) Then
            lblErro.Text = "Usuário ou senha inválidos!"
            lblErro.Visible = True
            txtSenha.Select()
            txtSenha.Focus()
            Exit Sub
        Else
            If Funcoes.VerificaPermissao(Processo, "LIBERAR", ObjUsuario.Usuario_Id) Then
                Dim par As New Dictionary(Of String, Object)
                par.Add("Usuario", ObjUsuario)
                par.Add("Processo", Processo)
                Session(Session("ssTipoRetorno")) = par
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(par)
                Else
                    Session("objSupervisor" & HID.Value) = par
                    CType(Me.Page, IBasePage).Carregar(par)
                End If
                Popup.CloseDialog(Me.Page, "divSupervisor")
            Else
                MsgBox(Me.Page, "Usuario Sem Permissão para Fechar o movimento Fiscal do Pedido")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Limpar()
        txtUsuario.Text = String.Empty
        txtSenha.Text = String.Empty
        lblErro.Text = String.Empty
    End Sub

    Protected Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Popup.CloseDialog(Me.Page, "divSupervisor")
    End Sub
End Class