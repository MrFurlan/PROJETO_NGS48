Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Data
Imports System.IO

Public Class AlteracaoEmpresaPedido
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        BuscaEmpresa()
        Limpar()
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub Limpar()
        txtClientes.Text = ""
        txtCodigoCliente.Value = ""
        Funcoes.VerificaEmpresa(ddlEmpresa)
        HID.Value = Guid.NewGuid().ToString
    End Sub

    Private Function ValidarCampos() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf txtClientes.Text.Length = 0 AndAlso txtPedido.Text.Length = 0 Then
            MsgBox(Me.Page, "Cliente não foi selecionado.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            If txtPedido.Text.Contains(",") OrElse txtPedido.Text.Contains(";") Then
                txtPedido.Text = txtPedido.Text.Replace(";", ",")
                For Each num As String In txtPedido.Text.Split(",")
                    If Not IsNumeric(num) Then
                        MsgBox(Me.Page, "Numero(s) de pedido inválido(s)")
                        Return False
                    End If
                Next
            Else
                If Not IsNumeric(txtPedido.Text) Then
                    MsgBox(Me.Page, "Numero de pedido inválido")
                    Return False
                End If
            End If
        End If
        Return True
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objClienteExtrato" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteExtrato" & HID.Value), [Lib].Negocio.Cliente))
                txtClientes.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteExtrato" & HID.Value)
            ElseIf Session("objPedidoExtrato" & HID.Value) IsNot Nothing Then
                Dim p As [Lib].Negocio.Pedido = CType(Session("objPedidoExtrato" & HID.Value), [Lib].Negocio.Pedido)
                txtPedido.Text = p.Codigo
                Session.Remove("objPedidoExtrato" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaCliente.Click
        Try
            Dim strJScript As String = ""
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteExtrato" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)
            HttpContext.Current.Session("txtCnpjDoCliente") = strCliente(0)
            HttpContext.Current.Session("txtEndDoCliente") = strCliente(1)

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedidoExtrato" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click

    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("ExtratoDePedido", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim CodCliente As String = ""
                    Dim EndCliente As Integer

                    If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                        CodCliente = txtCodigoCliente.Value.Split("-")(0)
                        EndCliente = txtCodigoCliente.Value.Split("-")(1)
                    End If

                    Dim strQueryString As String = ""

                    If ddlEmpresa.SelectedIndex > -1 Then strQueryString &= "&empresa=" & ddlEmpresa.SelectedValue
                    If txtClientes.Text.Length > 0 Then strQueryString &= "&cliente=" & txtCodigoCliente.Value.Replace(";", "-")
                    If txtPedido.Text.Length > 0 Then strQueryString &= "&pedido=" & txtPedido.Text

                    ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType, Guid.NewGuid().ToString, "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ExtratoDePedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class