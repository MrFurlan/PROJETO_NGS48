Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO

Public Class ucEmailNFePedido
    Inherits BaseUserControl

    Private objPedido As [Lib].Negocio.Pedido

    Public Property lstClientes() As [Lib].Negocio.ListCliente
        Get
            Return Session("_lstClientes")
        End Get
        Set(ByVal value As [Lib].Negocio.ListCliente)
            Session("_lstClientes") = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("_lstClientes")
        Session("ssNomeArquivoPedido" & HID.Value) = ""
        txtDestinatario.Text = ""
        txtAssunto.Text = ""
        txtMensagem.Text = ""
        grd.DataSource = New List(Of Object)
        grd.DataBind()
        txtDestinatario.Focus()
    End Sub

    Public Function EmailNFEPedido(ByVal parameters As Dictionary(Of String, Object)) As Integer
        txtDestinatario.Text = parameters("EmailNFEPedido").ToString

        objPedido = New Pedido(parameters("Empresa").ToString, parameters("EndEmpresa").ToString, parameters("Pedido").ToString)
        objPedido.ImprimirPedido(Me.Page, True, Session("ssNomeArquivoPedido" & HID.Value), True)

        Adicionar()
    End Function

    Private Sub Adicionar()
        If Not String.IsNullOrWhiteSpace(txtDestinatario.Text) Then
            If Not (lstClientes IsNot Nothing AndAlso lstClientes.Count > 0) Then
                lstClientes = New [Lib].Negocio.ListCliente
            End If

            Dim lst As String() = txtDestinatario.Text.Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each destinatario In lst
                Dim obj As New [Lib].Negocio.Cliente()
                obj.EmailNFE = destinatario.Trim()
                lstClientes.Add(obj)
            Next

            grd.DataSource = lstClientes
            grd.DataBind()
            txtDestinatario.Text = ""
            txtDestinatario.Focus()
        End If
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            If IsValid() Then
                Dim lstMail As New List(Of String)
                Dim smtp As New Net.Configuration.SmtpNetworkElement()
                Dim bodyHTML = txtMensagem.Text.Trim()

                Dim Sql As String = "Select * from Configuracao " & vbCrLf

                Dim objBanco As New AcessaBanco
                Dim dsMail As DataSet = objBanco.ConsultaDataSet(Sql, "ConfiguracaoXUsuario")
                If dsMail IsNot Nothing AndAlso dsMail.Tables IsNot Nothing AndAlso dsMail.Tables.Count > 0 AndAlso dsMail.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In dsMail.Tables(0).Rows
                        If Not String.IsNullOrWhiteSpace(row("Configuracao_Id")) Then
                            smtp.Host = row("Host")
                            smtp.Port = row("Porta")
                            smtp.Password = row("Senha")
                            smtp.UserName = row("Usuario")
                            Dim fromMail = row("Email")
                            Dim errorMsgMail As String = String.Empty
                            Dim subject As String = txtAssunto.Text.Trim()
                            Dim file = Session("ssNomeArquivoPedido" & HID.Value)

                            For Each obj As Cliente In lstClientes
                                lstMail.Add(obj.EmailNFE)
                            Next

                            If lstMail IsNot Nothing AndAlso lstMail.Count > 0 Then
                                If Funcoes.SendMail(fromMail, "NGS SOLUÇÕES", lstMail, subject, bodyHTML, smtp, errorMsgMail, file) Then
                                    MsgBox(Me.Page, "E-mail enviado com sucesso!", eTitulo.Info)
                                End If
                            End If
                        End If
                    Next
                End If

                Popup.CloseDialog(Me.Page, "divEmailNFePedido")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function IsValid()
        If Not (lstClientes IsNot Nothing AndAlso lstClientes.Count > 0) Then
            MsgBox(Me.Page, "É necessário possuir ao menos um e-mail de destinatário!")
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtAssunto.Text) Then
            MsgBox(Me.Page, "É necessário informar o campo assunto!")
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtMensagem.Text) Then
            MsgBox(Me.Page, "É necessário informar o campo mensagem!")
            Return False
        End If
        Return True
    End Function

    Protected Sub btnAdicionar_Click(sender As Object, e As EventArgs) Handles btnAdicionar.Click
        Adicionar()
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgExcluir As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(imgExcluir.NamingContainer, GridViewRow)
        lstClientes.RemoveAt(row.RowIndex)
        grd.DataSource = lstClientes
        grd.DataBind()
    End Sub

    Protected Sub lnkEnviar_Click(sender As Object, e As EventArgs) Handles lnkEnviar.Click
        Try
            Selecionar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divEmailNFePedido")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class