Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucOrigemDestino
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            ddl.Carregar(ddlViaTransporte, CarregarDDL.Tabela.ViaDeTransportes, "", False)
        End If
    End Sub

    Protected Sub btnOrigem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOrigem.Click
        Dim uc As ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)
        If uc IsNot Nothing Then
            uc.Limpar()
            uc.MainUserControl = Me
            Popup.ConsultaDeClientes(Me.Page, "objOrigemRoteiro" & HID.Value, "txtNome")
        End If
    End Sub

    Protected Sub btnDestino_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDestino.Click
        Dim uc As ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)
        If uc IsNot Nothing Then
            uc.Limpar()
            uc.MainUserControl = Me
            Popup.ConsultaDeClientes(Me.Page, "objDestinoRoteiro" & HID.Value, "txtNome")
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("objOrigemRoteiro" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
            hdfCodigoOrigem.Value = itemCliente.Value
            txtNomeOrigem.Text = itemCliente.Text
            Session.Remove("objOrigemRoteiro" & HID.Value)
        ElseIf Session("objDestinoRoteiro" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
            hdfCodigoDestino.Value = itemCliente.Value
            txtNomeDestino.Text = itemCliente.Text
            Session.Remove("objDestinoRoteiro" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Limpar()
        hdfCodigoOrigem.Value = ""
        hdfCodigoDestino.Value = ""
        txtNomeOrigem.Text = ""
        txtNomeDestino.Text = ""
    End Sub

    Protected Sub lnkConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConfirmar.Click
        If String.IsNullOrWhiteSpace(hdfCodigoOrigem.Value) OrElse String.IsNullOrWhiteSpace(hdfCodigoDestino.Value) Then
            MsgBox(Me.Page, "Origem e Destino são obrigatórios.")
        ElseIf hdfCodigoOrigem.Value = hdfCodigoDestino.Value Then
            MsgBox(Me.Page, "Origem e Destino não podem ser o mesmo.")
        Else
            Dim strCliente As String = String.Format("{0};{1};{2}", hdfCodigoOrigem.Value, hdfCodigoDestino.Value, ddlViaTransporte.SelectedValue)
            CType(Me.Page, IBasePage).Carregar(strCliente)
            Popup.CloseDialog(Me.Page, "divOrigemDestinoRoteiro")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        hdfCodigoOrigem.Value = String.Empty
        txtNomeOrigem.Text = String.Empty
        hdfCodigoDestino.Value = String.Empty
        txtNomeDestino.Text = String.Empty
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divOrigemDestinoRoteiro")
    End Sub

End Class