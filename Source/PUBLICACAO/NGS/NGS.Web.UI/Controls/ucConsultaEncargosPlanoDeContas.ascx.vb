Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaEncargosPlanoDeContas
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (IsPostBack) Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView(ByVal lst As ListEncargosPlanoDeContas)
        grd.DataSource = lst
        grd.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim lstSelecionados As List(Of Int32) = grd.GetSelectedItems("chkCodigo")
            Dim lstEncargosPlanoDeContas As ListEncargosPlanoDeContas = CType(Session("EncargosPlanoDeContas" & HID.Value), ListEncargosPlanoDeContas)
            For Each encargo As EncargosPlanoDeContas In lstEncargosPlanoDeContas
                For Each selecionado As String In lstSelecionados
                    If selecionado = encargo.CodigoContaEncargo Then
                        encargo.Selecionado = True
                        Continue For
                    Else
                        encargo.Selecionado = False
                    End If
                Next
            Next

            Session(Session("ssTipoRetorno")) = lstEncargosPlanoDeContas
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(lstEncargosPlanoDeContas)
                Else
                    CType(Me.Page, IBasePage).Carregar(lstEncargosPlanoDeContas)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaEncargosPlanoDeContas")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaEncargosPlanoDeContas")
    End Sub

    Protected Sub btnSelecionar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelecionar.Click
        Selecionar()
    End Sub

End Class