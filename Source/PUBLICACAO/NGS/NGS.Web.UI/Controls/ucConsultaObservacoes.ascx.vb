Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaObservacoes
    Inherits BaseUserControl

    Dim sql As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
        updConsultaObservacoes.Update()
    End Sub

    Public Sub BindGridView()
        Dim ds As New DataSet
        sql = "Select Codigo_Id, Estado , Descricao From ObservacoesTributarias order by Codigo_Id"
        ds = Banco.ConsultaDataSet(sql, "Observacoes")
        GridObservacoes.DataSource = ds
        GridObservacoes.DataBind()
        updConsultaObservacoes.Update()
    End Sub

    Protected Sub GridObservacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridObservacoes.SelectedIndexChanged
        Try
            Dim Observacoes As String = Session("Observacoes" & HID.Value.ToString)
            If Observacoes.Length > 0 Then
                Observacoes = Observacoes & " " & GridObservacoes.SelectedRow.Cells(3).Text()
            Else
                Observacoes = GridObservacoes.SelectedRow.Cells(3).Text()
            End If

            HttpContext.Current.Session("Observacoes" & HID.Value.ToString) = Observacoes

            If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objObservacoesNXI")) Then
                Dim txtObservacoesFiscais = CType(WebHelpers.FindControlRecursive(Me.Page, "txtObservacoesFiscais"), TextBox)
                If Not txtObservacoesFiscais Is Nothing Then
                    txtObservacoesFiscais.Text = Observacoes
                End If
                Popup.CloseDialog(Me.Page, "divConsultaObservacoes")
            ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objObservacaoAxE")) Then
                Dim txtObservacao = CType(WebHelpers.FindControlRecursive(Me.Page, "txtObservacao"), TextBox)
                If Not txtObservacao Is Nothing Then
                    txtObservacao.Text = Observacoes
                End If
                Popup.CloseDialog(Me.Page, "divConsultaObservacoes")
            ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objObservacaoEmbarquexObs")) Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Session(HttpContext.Current.Session("ssTipoRetorno")) = Observacoes
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(Observacoes)
                Else
                    CType(Me.Page, IBasePage).Carregar(Observacoes)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaObservacoes")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divConsultaObservacoes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

End Class