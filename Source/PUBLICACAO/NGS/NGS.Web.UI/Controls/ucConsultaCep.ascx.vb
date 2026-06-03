Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaCep
    Inherits BaseUserControl

    Public Property CepSelecionado() As GridViewRow
        Get
            Return CType(ViewState("dsCep"), GridViewRow)
        End Get
        Set(value As GridViewRow)
            ViewState("dsCep") = value
        End Set
    End Property

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Me.Limpar()
            ddl.Carregar(ddlEstado, CarregarDDL.Tabela.EstadoERegiao, "", True)
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If String.IsNullOrWhiteSpace(txtCidade.Text) Then
                MsgBox(Me.Page, "Informe a cidade")
            ElseIf ddlEstado.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Informe o Estado")
            ElseIf String.IsNullOrWhiteSpace(txtRua.Text) Then
                MsgBox(Me.Page, "Informe no minimo 3 caracteres para o endereço.")
            Else
                Dim ds As DataSet = Funcoes.ConsultaCep(ddlEstado.SelectedValue, txtCidade.Text, txtRua.Text)
                gridCep.DataSource = ds.Tables(1)
                gridCep.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaCep")
    End Sub

    Private Function Decode(ByVal str As String) As String
        Return Server.HtmlDecode(str)
    End Function

    Protected Sub gridCep_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridCep.SelectedIndexChanged
        Try
            Try

                Session(Session("ssTipoRetorno")) = gridCep.Rows(0)
                If Session("ssTipoRetorno") IsNot Nothing Then
                    Dim str As String = String.Format("{0}-{1}-{2}-{3}-{4}-{5}", Decode(gridCep.SelectedRow.Cells(1).Text).Replace("-", ""), Decode(gridCep.SelectedRow.Cells(2).Text), _
                                                                                 Decode(gridCep.SelectedRow.Cells(3).Text), Decode(gridCep.SelectedRow.Cells(4).Text), _
                                                                                 Decode(gridCep.SelectedRow.Cells(5).Text), Decode(gridCep.SelectedRow.Cells(6).Text))
                    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                        Dim ucName = MainUserControl.ClientID.Split("_")
                        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                        CType(uc, IBaseUserControl).Carregar(str)
                    Else
                        CType(Me.Page, IBasePage).Carregar(str)
                    End If
                    Popup.CloseDialog(Me.Page, "divConsultaCep")
                End If
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Overrides Sub Limpar()
        txtCidade.Text = String.Empty
        'txtEstado.Text = String.Empty
        ddlEstado.SelectedIndex = 0
        txtRua.Text = String.Empty
    End Sub
End Class