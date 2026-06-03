Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucEmailNFe
    Inherits BaseUserControl

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
        txtDestinatario.Text = ""
        txtAssunto.Text = ""
        txtMensagem.Text = ""
        grd.DataSource = New List(Of Object)
        grd.DataBind()
        txtDestinatario.Focus()
    End Sub

    Public Function EmailNFE(ByVal parameters As Dictionary(Of String, Object)) As Integer
        txtDestinatario.Text = parameters("EmailNFE").ToString
        txtAssunto.Text = "FATURAMENTO - " & parameters("Cliente").ToString & " - " & parameters("Pedido").ToString & " - " & parameters("Placa").ToString
        txtMensagem.Text = "Prezados Senhores," & vbCrLf & vbCrLf &
                           "Segue em anexo a nossa nota fiscal eletrônica, conforme indicado abaixo:" & vbCrLf & vbCrLf &
                           "Nota Fiscal: " & parameters("NF").ToString & "" & vbCrLf &
                           "Emitida em: " & parameters("Emissao").ToString & vbCrLf &
                           "Produto: " & parameters("Produto").ToString & vbCrLf &
                           "Quantidade: " & parameters("Quantidade").ToString & vbCrLf &
                           "Valor Unitário: " & parameters("ValorUnitario") & vbCrLf &
                           "Valor Total: " & parameters("ValorTotal") & vbCrLf &
                           "Chave Acesso NFE: " & parameters("ChaveNFE") & vbCrLf &
                           "Emitente: " & parameters("Empresa") & vbCrLf & vbCrLf &
                           "Adicionalmente a nota fiscal pode ser consultada seguintes sites: " & vbCrLf &
                           "Ambiente nacional: http://www.nfe.fazenda.gov.br" & vbCrLf & vbCrLf &
                           "Atenciosamente," & vbCrLf & parameters("Empresa") & vbCrLf

        Adicionar()
    End Function

    Public Function EnvioDeEmail(ByVal parameters As Dictionary(Of String, Object)) As Integer
        txtAssunto.Text = parameters("Assunto").ToString
        txtMensagem.Text = parameters("Mensagem").ToString

        If parameters("Compactado").ToString() = "SIM" Then
            chkCompactado.Checked = True
        Else
            chkCompactado.Checked = False
        End If
    End Function

    Protected Overrides Sub Selecionar()
        Try
            If Not IsValid() Then
                Exit Sub
            End If
            Session("strAssunto" & HID.Value) = txtAssunto.Text
            Session("strMensagem" & HID.Value) = txtMensagem.Text
            Session("strCompactado" & HID.Value) = chkCompactado.Checked
            Session(Session("ssTipoRetorno")) = lstClientes
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(lstClientes)
                Else
                    CType(Me.Page, IBasePage).Carregar(lstClientes)
                End If
                Popup.CloseDialog(Me.Page, "divEmailNFe")
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
            Popup.CloseDialog(Me.Page, "divEmailNFe")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class