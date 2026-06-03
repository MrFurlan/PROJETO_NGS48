Imports System.Runtime.InteropServices.ComTypes
Imports NGS.Lib.Negocio

Public Class PedidoOnsoft
    Inherits BasePage

#Region "Barra Botão"

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            CarregarPedidosOnMobile()
            CarregarPedidoNAOIntegrado()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Limpar()
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

        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("PedidoOnsoft", "ACESSAR") Then
                BuncarUnidadeDeNegocio()
                txtDataInicial.Text = Now.AddDays(-30).ToString("dd/MM/yyyy")
                txtDataFinal.Text = Now.ToString("dd/MM/yyyy")
                Limpar()
                CarregarPedidoNAOIntegrado()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If

    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClientePedidoOnSoft" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Private Sub Limpar()
        txtPedido.Text = ""
        gridPendente.DataSource = Nothing
        gridPendente.DataBind()
    End Sub

    Private Sub CarregarPedidoNAOIntegrado()
        Dim sql As String = "SELECT pni.VendedorCod, cv.Nome as NomeVendedor, pni.PedidoNum, pni.ClienteCod, c.Nome as NomeCliente, " & vbCrLf &
                            "		pni.TabPrecoCod, pni.PedidoData, pni.PedidoVrpagar" & vbCrLf &
                            "FROM PedidoNAOIntegracaoOnSoft pni" & vbCrLf &
                            "		left join Clientes cv" & vbCrLf &
                            "				on cv.Cliente_Id   = pni.VendedorCod" & vbCrLf &
                            "				and cv.Endereco_Id = 0" & vbCrLf &
                            "		left join Clientes c" & vbCrLf &
                            "				on c.Cliente_Id   = pni.ClienteCod" & vbCrLf &
                            "				and c.Endereco_Id = 0"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "onMobile")

        If ds.Tables(0).Rows.Count > 0 Then
            gridPendente.DataSource = ds
        Else
            gridPendente.DataSource = Nothing
        End If

        gridPendente.DataBind()

    End Sub
    Private Sub CarregarPedidosOnMobile()

        If (String.IsNullOrWhiteSpace(txtDataInicial.Text)) Or String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data início e Data final são obrigatórios")
            Return
        End If

        Dim clienteId = ""

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Dim cliente = txtCodigoCliente.Value.ToString.Split("-")
            clienteId = cliente(0)
        End If

        Dim tSituacao As String = String.Empty

        If rdBloqueado.Checked Then
            tSituacao = "B"
        ElseIf rdLiberado.Checked Then
            tSituacao = "L"
        End If

        Dim pedidos As New ListPedidoOnMobile(txtDataInicial.Text, txtDataFinal.Text, clienteId, IIf(String.IsNullOrWhiteSpace(txtPedido.Text), 0, txtPedido.Text), tSituacao)

        gridPedidoOnsoft.DataSource = pedidos
        gridPedidoOnsoft.DataBind()

        Dim i As Integer = 0
        While i < gridPedidoOnsoft.Rows.Count
            If CType(gridPedidoOnsoft.Rows(i).FindControl("hidPedidoLiberado"), HiddenField).Value Then
                CType(gridPedidoOnsoft.Rows(i).FindControl("imgLiberadoBloqueado"), ImageButton).ImageUrl = "~/Images/erro.jpg"
                CType(gridPedidoOnsoft.Rows(i).FindControl("imgLiberadoBloqueado"), ImageButton).ToolTip = "Bloqueado"
            Else
                CType(gridPedidoOnsoft.Rows(i).FindControl("imgLiberadoBloqueado"), ImageButton).ImageUrl = "~/Images/certo.jpg"
                CType(gridPedidoOnsoft.Rows(i).FindControl("imgLiberadoBloqueado"), ImageButton).ToolTip = "Liberado"
            End If
            i += 1
        End While

    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Session("objClientePedidoOnSoft" & HID.Value) IsNot Nothing Then
                Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClientePedidoOnSoft" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
                txtClientes.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClientePedidoOnSoft" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

End Class