Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaClientesDireto
    Inherits BaseUserControl

    Dim Clientes As New [Lib].Negocio.ListCliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
        Limpar()
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            Dim cliente As String() = args.Split(";")
            Dim objCliente As New [Lib].Negocio.Cliente(cliente(0), Convert.ToInt32(cliente(1)))
            Session(Session("ssTipoRetorno")) = objCliente
            If String.IsNullOrWhiteSpace(objCliente.CEP) Then
                MsgBox(Me.Page, "Verifique o CEP informado no cadastro do cliente.", eTitulo.Info)
            Else
                If Session("ssTipoRetorno") IsNot Nothing Then
                    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                        CType(uc, IBaseUserControl).Carregar(objCliente)
                    Else
                        CType(Me.Page, IBasePage).Carregar(objCliente)
                    End If
                    Popup.CloseDialog(Me.Page, "divConsultaClienteDireto")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gdvClientes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gdvClientes.SelectedIndexChanged
        Dim strCodigo As String = gdvClientes.SelectedRow.Cells(5).Text().Replace(".", "").Replace("/", "").Replace("-", "")
        Dim strEndereco As String = gdvClientes.SelectedRow.Cells(6).Text()
        Selecionar(strCodigo & ";" & strEndereco)
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.DestroyDialog(Me.Page, "divConsultaClienteDireto")
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("_MainUserControl")
        gdvClientes.DataSource = New List(Of Object)()
        gdvClientes.DataBind()
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SetarTituloDIV(ByVal Titulo As String)
        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "$(document).ready(function () { $('#divConsultaClienteDireto').prop('title', '" & Titulo & "'); });", True)
    End Sub

    Public Sub BindGridView(ByVal lst As List(Of [Lib].Negocio.Cliente))
        gdvClientes.DataSource = lst
        gdvClientes.DataBind()
        If (gdvClientes.Rows.Count > 0) Then
            gdvClientes.Rows(0).Cells(0).Focus()
        End If
    End Sub

    Public Sub CarregarTransportadoresPorPedido(ByRef Transportadores As [Lib].Negocio.ListPedidoXTransportador)
        If (Not Transportadores Is Nothing) Then
            Dim cliente As New [Lib].Negocio.Cliente
            For i = 0 To Transportadores.Count - 1
                cliente = CType(Transportadores(i).Transportador, [Lib].Negocio.Cliente)
                Clientes.Add(cliente)
            Next
            gdvClientes.DataSource = Clientes
            gdvClientes.DataBind()
            If (gdvClientes.Rows.Count > 0) Then
                gdvClientes.Rows(0).Cells(0).Focus()
            End If
        End If
    End Sub

    Public Sub CarregarDepositosPorPedido(ByRef Depositos As [Lib].Negocio.ListPedidoxDeposito)
        If (Not Depositos Is Nothing) Then
            Dim deposito As New [Lib].Negocio.Cliente
            For i = 0 To Depositos.Count - 1
                deposito = CType(Depositos(i).Deposito, [Lib].Negocio.Cliente)
                Clientes.Add(deposito)
            Next
            gdvClientes.DataSource = Clientes
            gdvClientes.DataBind()
            If (gdvClientes.Rows.Count > 0) Then
                gdvClientes.Rows(0).Cells(0).Focus()
            End If
        End If
    End Sub

End Class