Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Expedicao
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)

        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("Expedicao", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar esse módulo!", IIf(Not String.IsNullOrEmpty(Session("ssMenuDeAcesso")), Session("ssMenuDeAcesso") & ".aspx", "~/Index.aspx"), eTitulo.Info)
                Exit Sub
            End If

            If UsuarioServidor.Usuario.Empresa.Empresa.ConferenciaNFE AndAlso Funcoes.VerificaPermissao("Fiscal", "ACESSAR") Then
                PesquisaNotasConferencia()
            End If

        End If
    End Sub

    Protected Sub PesquisaNotasConferencia()
        Dim NF As New [Lib].Negocio.NotaFiscal
        gdvConferenciaNota.DataSource = NF.ConferenciaNotas()
        gdvConferenciaNota.DataBind()
        tblNotasConferencia.Visible = gdvConferenciaNota.Rows.Count > 0
    End Sub

    Protected Sub lnkConferir_Click(sender As Object, e As EventArgs)
        Dim lnkConferir As LinkButton = CType(sender, LinkButton)
        Dim row As GridViewRow = CType(lnkConferir.NamingContainer, GridViewRow)
        Dim param As New Dictionary(Of String, Object)
        param.Add("Nota", row.Cells(0).Text)
        param.Add("Serie", row.Cells(1).Text)
        param.Add("CodigoEmpresa", row.Cells(2).Text)
        param.Add("EnderecoEmpresa", row.Cells(3).Text)
        param.Add("CodigoCliente", row.Cells(5).Text)
        param.Add("EnderecoCliente", row.Cells(6).Text)
        Session("ssConferenciaNFe") = param
        Response.Redirect("NotasFiscaisGerais.aspx")
    End Sub

End Class