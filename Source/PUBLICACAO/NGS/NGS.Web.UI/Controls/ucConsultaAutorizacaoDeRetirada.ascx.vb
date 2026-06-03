Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaAutorizacaoDeRetirada
    Inherits BaseUserControl

    Dim objNotaFiscal As NotaFiscal

    Public Property MyParameters() As Dictionary(Of String, Object)
        Get
            Return CType(Session("MyParameters" & HID.Value), Dictionary(Of String, Object))
        End Get
        Set(ByVal value As Dictionary(Of String, Object))
            Session("MyParameters" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Dim Tipo As String = Session("ssTipoRetorno")
        If Session(Tipo) Is Nothing Then
            Session("Sem" & Tipo) = -1
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack Then
            Return
        End If
        txtAno.Text = DateTime.Now.Year.ToString
    End Sub

    Protected Sub gridAutorizacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Tipo As String = Session("ssTipoRetorno")

        'Edson *****
        Dim Par As New Hashtable
        Par.Add("Pedido", gridAutorizacao.SelectedRow.Cells(2).Text)
        Par.Add("Autorizacao", gridAutorizacao.SelectedRow.Cells(1).Text)
        '***********

        If CDbl(gridAutorizacao.SelectedRow.Cells(7).Text) = 0 AndAlso CDbl(gridAutorizacao.SelectedRow.Cells(10).Text) = 0 Then
            Session("SemAutorizacao" & HID.Value) = -1
        Else
            Session(Tipo) = gridAutorizacao.SelectedRow.Cells(1).Text
        End If

        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objAutorizacaoNXI")) Then
            CType(Me.Page, NotaFiscalXItens).CarregarAutorizacao(Par)
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objAutorizacaoAxE")) Then
            CType(Me.Page, AutorizacaoDeRetirada).CarregarAutorizacaoAxE(Par)
            Session.Remove(Tipo)
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objLaudo")) Then
            CType(Me.Page, Laudo).CarregarAutorizacao(Par)
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objAutorizacaoPesagem")) Then
            CType(Me.Page, AlterarLaudo).CarregarAutorizacao(Par)
            Session.Remove(Tipo)
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objRateioDePesagem")) Then
            CType(Me.Page, RateioDePesagem).CarregarAutorizacao(Par)
            Session.Remove(Tipo)
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objSelRateioDePesagem")) Then
            CType(Me.Page, RateioDePesagem).CarregarAutorizacao(Par)
            Session.Remove(Tipo)
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objAutorizacaoRxP")) Then
            Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
            CType(uc, ucRateioDePesagem).CarregarAutorizacao(Par)
            Session.Remove(Tipo)
        End If

        Popup.CloseDialog(Me.Page, "divConsultaAutorizacaoDeRetirada")
    End Sub
    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaAutorizacaoDeRetirada")
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView(Optional ByVal parameters As Dictionary(Of String, Object) = Nothing)
        SessaoRecuperaNotaFiscal()
        If (objNotaFiscal IsNot Nothing) Then
            Dim Tipo As String = "Autorizacao" & HID.Value
            Dim Ped As String = objNotaFiscal.CodigoPedido
            Dim Emp As String = objNotaFiscal.CodigoEmpresa
            Dim EndEmp As String = objNotaFiscal.EnderecoEmpresa
            Dim cli As String = ""
            Dim Endcli As String = ""
            Dim Romaneio As Boolean = False

            Dim Autorizacoes As New ListAutorizacaoDeRetirada(Emp, EndEmp, cli, IIf(Endcli.Length = 0, 0, Endcli), Ped, objNotaFiscal.SubOperacao.Classe, Romaneio, txtAno.Text)
            gridAutorizacao.DataSource = Autorizacoes
            gridAutorizacao.DataBind()

            'Edson *****
            Dim Par As New Hashtable
            Par.Add("Pedido", "")
            Par.Add("Autorizacao", 0)
            '***********

            If Autorizacoes Is Nothing OrElse Autorizacoes.Count = 0 Then
                If objNotaFiscal.CodigoAutorizacao = -1 Then MsgBox(Me.Page, "Autorização de Retirada não foi encontrada.")

                Session("Sem" & Tipo) = -1
                If TypeOf Me.Page Is NotaFiscalXItens Then
                    CType(Me.Page, NotaFiscalXItens).CarregarAutorizacao(Par)
                ElseIf TypeOf Me.Page Is RateioDePesagem Then
                    CType(Me.Page, RateioDePesagem).CarregarAutorizacao(Par)
                End If
            End If
        ElseIf (parameters IsNot Nothing) Then
            Dim Tipo As String = "Autorizacao" & HID.Value
            Dim Ped As String = parameters("ped")
            Dim Emp As String = parameters("emp")
            Dim EndEmp As String = parameters("endemp")
            Dim cli As String = parameters("cli")
            Dim Endcli As String = parameters("endcli")
            Dim classe As [Lib].Negocio.eClassesOperacoes = parameters("classe")
            Dim Romaneio As Boolean = False
            If (parameters.ContainsKey("romaneio") AndAlso parameters("romaneio") IsNot Nothing) Then
                Romaneio = CBool(parameters("romaneio"))
            End If

            MyParameters = parameters
            Dim Autorizacoes As New ListAutorizacaoDeRetirada(Emp, EndEmp, cli, IIf(Endcli.Length = 0, 0, Endcli), Ped, classe, Romaneio, txtAno.Text)
            gridAutorizacao.DataSource = Autorizacoes
            gridAutorizacao.DataBind()

            'Edson *****
            Dim Par As New Hashtable
            Par.Add("Pedido", "")
            Par.Add("Autorizacao", 0)
            '***********

            If Autorizacoes Is Nothing OrElse Autorizacoes.Count = 0 Then
                Session("Sem" & Tipo) = -1
                If TypeOf Me.Page Is NotaFiscalXItens Then
                    CType(Me.Page, NotaFiscalXItens).CarregarAutorizacao(Par)
                ElseIf TypeOf Me.Page Is RateioDePesagem Then
                    CType(Me.Page, RateioDePesagem).CarregarAutorizacao(Par)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaAutorizacaoDeRetirada")
            End If
        End If
    End Sub

    Protected Sub btnBuscar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnBuscar.Click
        BindGridView(MyParameters)
    End Sub

End Class