Imports NGS.Lib.Negocio

Public Class ucConsultaObservacoesEmbarque
    Inherits BaseUserControl

    Private objEmbarquePedido As [Lib].Negocio.EmbarquePedido

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Private Sub SessaoSalvarObjEmbarque()
        Session("ssEmbarquePedido" & HID.Value) = objEmbarquePedido
    End Sub

    Private Sub SessaoRecuperaObjEmbarque()
        objEmbarquePedido = CType(Session("ssEmbarquePedido" & HID.Value), [Lib].Negocio.EmbarquePedido)
    End Sub


    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Protected Sub lnkSair_Click(sender As Object, e As EventArgs) Handles lnkSair.Click
        SessaoRecuperaObjEmbarque()
        objEmbarquePedido.LocaisDeEntrega(hdLocalDeEntrega.Value).Observacao = Funcoes.EliminarCaracteresEspeciais(txtObservacaoEmbarque.Text)
        objEmbarquePedido.LocaisDeEntrega(hdLocalDeEntrega.Value).IUD = "U"
        If Not objEmbarquePedido.LocaisDeEntrega(hdLocalDeEntrega.Value).Salvar() Then
            MsgBox(Me.Page, "Erro ao salvar a mensagem")
        End If
        SessaoSalvarObjEmbarque()
        Popup.CloseDialog(Me.Page, "divConsultaObservacoesEmbarque")
    End Sub

    Public Sub SetarLocalEntregaProduto(ByVal linhaLocalEntrega As Integer)
        hdLocalDeEntrega.Value = linhaLocalEntrega
        SessaoRecuperaObjEmbarque()
        txtObservacaoEmbarque.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Observacao
    End Sub

    Protected Sub lnkConsultarObservacao_Click(sender As Object, e As EventArgs) Handles lnkConsultarObservacao.Click
        Session("Observacoes" & HID.Value) = txtObservacaoEmbarque.Text

        Dim ucConsultaObservacoes As ucConsultaObservacoes = DirectCast(Me.NamingContainer.FindControl("ucConsultaObservacoes"), ucConsultaObservacoes)
        If ucConsultaObservacoes IsNot Nothing Then
            'ucConsultaObservacoes.Limpar()
            ucConsultaObservacoes.MainUserControl = Me
            ucConsultaObservacoes.BindGridView()
            Popup.ConsultaDeObservacoes(Me.Page, "objObservacaoEmbarquexObs" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Carregar(str As String)
        If Session("objObservacaoEmbarquexObs" & HID.Value) IsNot Nothing Then
            txtObservacaoEmbarque.Text = str
            Session.Remove("objObservacaoEmbarquexObs" & HID.Value)
        End If
    End Sub

End Class