Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaFaturasDeFrete
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("ssFaturasCons" & HID.Value)
        Session.Remove("_MainUserControl")
        DgFaturas.DataSource = New List(Of Object)
        DgFaturas.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim lstFaturas As [Lib].Negocio.ListFaturaDeFrete = Session("ssFaturasCons" & HID.Value)
            Dim objFaturaDeFrete = lstFaturas(DgFaturas.SelectedIndex)
            Session(Session("ssTipoRetorno")) = objFaturaDeFrete

            If objFaturaDeFrete.ListTituloFatura IsNot Nothing Then
                For Each t In objFaturaDeFrete.ListTituloFatura
                    If t.Titulo.CodigoProvisao = eProvisao.Baixa Then
                        MsgBox(Me.Page, "O título " & t.Titulo.Codigo & " está baixado. Caso ainda tenha saldo para a fatura ajuste pelo Financeiro.")
                        Exit Sub
                    End If
                Next
            End If

            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objFaturaDeFrete)
                Else
                    CType(Me.Page, IBasePage).Carregar(objFaturaDeFrete)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaFaturasDeFrete")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub




    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))
        Dim Num As String = parameters("Num")
        Dim Emp As String = parameters("Emp")
        Dim EndEmp As String = parameters("EndEmp")
        Dim Conv As String = parameters("Conv")
        Dim EndConv As String = parameters("EndConv")

        Dim objFatura As New [Lib].Negocio.FaturaDeFrete()
        objFatura.CodigoFatura = Num
        objFatura.CodigoEmpresa = Emp
        objFatura.EnderecoEmpresa = EndEmp
        objFatura.CodigoConveniado = Conv
        objFatura.EnderecoConveniado = EndConv
        Dim lstFaturas As New [Lib].Negocio.ListFaturaDeFrete(objFatura, False, New Cliente, String.Empty, String.Empty, IIf(FinanceiroNovo, "2", "3"))

        If lstFaturas IsNot Nothing AndAlso lstFaturas.Count > 0 Then
            Session("ssFaturasCons" & HID.Value) = lstFaturas
            DgFaturas.DataSource = lstFaturas
            DgFaturas.DataBind()
        Else
            DgFaturas.DataBind()
            MsgBox(Me.Page, "Registro(s) não encontrado(s) para esta seleção/período!")
            Popup.CloseDialog(Me.Page, "divConsultaFaturasDeFrete")
        End If
    End Sub

    Protected Sub DgFaturas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaFaturasDeFrete")
    End Sub

End Class