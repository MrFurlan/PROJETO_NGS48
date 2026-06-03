Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class MetasVendas
    Inherits BasePage

    Dim ObjListMetasXSafras As New [Lib].Negocio.ListMetasXSafras
    Dim ObjMetasXSafras As New [Lib].Negocio.MetasXSafras
    Dim ObjMetas As New [Lib].Negocio.Metas

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("MetasVendas", "ACESSAR") Then
                ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2009;5;I")
                ddl.Carregar(ddlAno2, CarregarDDL.Tabela.Ano, "2009;5;I")
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlCultura, CarregarDDL.Tabela.Cultura, "")
                HID.Value = Guid.NewGuid().ToString
                ucConsultaEmpresas.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
            End If
        End If
    End Sub

    Public Sub CarregarEmpresa()
        Try
            If Not Session("ObjEmpresaMV" & HID.Value) Is Nothing Then
                Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("ObjEmpresaMV" & HID.Value), [Lib].Negocio.Cliente))
                txtEmpresa.Text = itemEmpresa.Text
                txtEmpresa2.Text = itemEmpresa.Text
                txtCodigoEmpresa.Value = itemEmpresa.Value
                Session.Remove("ObjEmpresaMV" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getSql()
        Dim sql As String = "SELECT        M.Empresa, MXS.Safra_Id, MXS.Cultura_Id" & vbCrLf & _
                            "   FROM            Metas M " & vbCrLf & _
                            "   INNER JOIN  MetasXSafras MXS" & vbCrLf & _
                            "        ON M.Meta_Id = MXS.Meta_Id" & vbCrLf

        Return sql
    End Function

    Protected Sub btnConsultaEmpresa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultaEmpresa.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "ObjEmpresaMV" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtEmpresa_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtEmpresa.TextChanged
        Dim ds As DataSet

        Try
            If String.IsNullOrWhiteSpace(txtEmpresa.Text) Then
                ds = Banco.ConsultaDataSet(getSql, "Consulta")
                GridMetas.DataSource = ds
                GridMetas.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnNovo.Click
        Try
            If Funcoes.VerificaPermissao("MetasVendas", "GRAVAR") Then
                ddlAno2.SelectedIndex = ddlAno.SelectedIndex
                txtEmpresa2.Text = txtEmpresa.Text
                TabContainer1.ActiveTabIndex = 1
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaEmpresa2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultaEmpresa2.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "ObjEmpresaMV" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAdicionar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdicionar.Click
        Dim arrMetas As String()

        Try
            arrMetas = txtCodigoEmpresa.Value.Split("-")
            ObjMetas.Empresa = arrMetas(0)
            ObjMetas.Ano = ddlAno2.SelectedValue
            ObjMetas.CotacaoDollar = CDec("2.20")
            ObjMetas.IUD = "I"

            If ddlSafra.SelectedIndex > 0 Then
                ObjMetasXSafras.CodigoSafra = ddlSafra.SelectedValue
            End If
            If ddlCultura.SelectedIndex > 0 Then
                ObjMetasXSafras.CodigoCultura = ddlCultura.SelectedValue
            End If
            ObjMetas.Salvar()
            ObjMetasXSafras.Salvar()
            MsgBox(Me.Page, "Metas inserida com Sucesso.", eTitulo.Sucess)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "MetasVendas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub AjudaMeta2_Click(sender As Object, e As EventArgs) Handles AjudaMeta2.Click
        Try
            Funcoes.Ajuda(Me.Page, "MetasVendas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class