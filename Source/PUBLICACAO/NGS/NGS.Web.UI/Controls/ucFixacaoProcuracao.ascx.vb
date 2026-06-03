Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucFixacaoProcuracao
    Inherits BaseUserControl
    
#Region "Variaveis Locais"
    Private objPedido As [Lib].Negocio.Pedido
#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkCessaoDeCredito_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)
            Session(Session("ssTipoRetorno")) = row.Cells(1).Text
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(row.Cells(1).Text, row.Cells(5).Text)
                Else
                    CType(Me.Page, IBasePage).Carregar(row.Cells(1).Text, row.Cells(5).Text)
                End If
                Popup.CloseDialog(Me.Page, "divFixacaoProcuracao")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divFixacaoProcuracao")
    End Sub
#End Region

#Region "Methods"
    Private Sub SessaoRecuperaPedido()
        objPedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
        If objPedido Is Nothing Then objPedido = New [Lib].Negocio.Pedido
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Function CarregarProcuracoes(ByVal itemProduto As Integer) As Boolean
        SessaoRecuperaPedido()
        grdProcuracoes.DataSource = objPedido.Itens(itemProduto).Fixacoes.ListarProcuracao()
        grdProcuracoes.DataBind()
    End Function
#End Region
End Class