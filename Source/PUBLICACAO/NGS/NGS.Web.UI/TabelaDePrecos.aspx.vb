Imports System.Data
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TabelaDePrecos
    Inherits BasePage

    Dim objCliXPro As ClienteXProduto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsConnect AndAlso Not IsPostBack Then
            ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
            ddl.Carregar(ddlGruposProdutos, CarregarDDL.Tabela.GrupoProdutoXConsumo)
            ddl.Carregar(ddlEncargos, CarregarDDL.Tabela.Encargos, "Encargo_id in ('PIS')")
            Limpar()
        End If
    End Sub

#Region "Metodos"
    Private Sub carregarGrid()
        Dim cliente As String() = txtCliente.Text.Trim.Split("-")
        Dim objListCliXPro = New ListClienteXProduto(cliente(0), cliente(1))

        gridClienteXProdutos.DataSource = objListCliXPro
        gridClienteXProdutos.DataBind()

        Session("objListClienteXProduto" & HID.Value) = objListCliXPro
    End Sub

    Private Sub Limpar()
        Session.Remove("objTblPrecos" & HID.Value)
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objProduto" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
    End Sub
#End Region

#Region "Eventos"
    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProdutos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProdutos.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProduto" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGruposProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlGruposProdutos.SelectedIndexChanged
        Try
            ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.ProdutoProducao, " P.Grupo = '" & ddlProdutos.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsulta_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultar.Click
        Try
            carregarGrid()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

End Class