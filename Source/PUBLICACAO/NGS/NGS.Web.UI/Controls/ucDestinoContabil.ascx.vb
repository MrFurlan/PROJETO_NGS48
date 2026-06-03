Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucDestinoContabil
    Inherits BaseUserControl

    Private sql As String
    Private strJavaScript As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtClienteDestino.Text = ""
        txtCodigoClienteDestino.Value = ""
        ddlCarteiraDestino.SelectedIndex = 0
    End Sub

    Public Overrides Sub Carregar(ByVal parameters As Dictionary(Of String, Object))
        sql = "SELECT  Produto_Id AS Codigo, Produto_Id + '  -  ' + Descricao AS Descricao  FROM ComprasXProdutos Where Classificacao = '" & parameters("tipo") & "' Order By Produto_Id"
        ddlCarteiraDestino.DataValueField = "Codigo"
        ddlCarteiraDestino.DataTextField = "Descricao"
        ddlCarteiraDestino.DataSource = Banco.ConsultaDataSet(sql, "Carteiras")
        ddlCarteiraDestino.DataBind()
        Funcoes.InserirLinhaEmBranco(ddlCarteiraDestino)
        ddlCarteiraDestino.SelectedIndex = 0
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteDESXCONT" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
            txtClienteDestino.Text = itemCliente.Text
            txtCodigoClienteDestino.Value = itemCliente.Value
            Session.Remove("objClienteDESXCONT" & HID.Value)
            Session("ssTipoRetorno") = Session("ssCampo" & HID.Value)
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo" & HID.Value) = Session("ssTipoRetorno")
        If TypeOf Me.Page Is WFTitulo Then
            Dim ucConsultaClientes As ucConsultaClientes = CType(WebHelpers.FindControlRecursive(Me.Page, "ucConsultaClientes"), ucConsultaClientes)
            If ucConsultaClientes IsNot Nothing Then
                ucConsultaClientes.Limpar()
                ucConsultaClientes.MainUserControl = Me
                Popup.ConsultaDeClientes(Me.Page, "objClienteDESXCONT" & HID.Value)
            End If
        End If
    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If String.IsNullOrWhiteSpace(txtCodigoClienteDestino.Value) Then
            MsgBox(Me.Page, "Cliente não foi selecionado!")
        ElseIf ddlCarteiraDestino.SelectedIndex = -1 Then
            MsgBox(Me.Page, "Finalidade financeira não foi selecionada!")
        ElseIf ddlCarteiraDestino.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Finalidade financeira não foi selecionada!")
        Else
            Dim Tipo As String = Session("ssTipoRetorno")
            Session(Tipo) = txtCodigoClienteDestino.Value & "-" & ddlCarteiraDestino.SelectedValue
            Popup.CloseDialog(Me.Page, "divDestinoContabil")
        End If
    End Sub

    Protected Sub btnSair_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Popup.CloseDialog(Me.Page, "divDestinoContabil")
    End Sub

End Class