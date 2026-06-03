Imports System.Data
Imports System
Imports System.Globalization
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BalancoAuditadoMes
    Inherits BasePage

    Dim objListBalanco As ListBalancoAuditadoMes

#Region "Session"
    Private Sub SessaoSalvaBalanco()
        Session("objListBalanco" & Hid.Value) = objListBalanco
    End Sub

    Private Sub SessaoRecuperaBalanco()
        objListBalanco = CType(Session("objListBalanco" & Hid.Value), [Lib].Negocio.ListBalancoAuditadoMes)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("BALANCOAUDITADOMES", "ACESSAR") Then
                ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2016;10;I", True)
                ddlAno.SelectedValue = IIf(Now.Month = 1, Now.Year - 1, Now.Year)
                CarregarMovimento()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
            End If
        End If
    End Sub

    Protected Sub gridMeses_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridMeses.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim btnJaneiro As Button = e.Row.FindControl("btnJaneiro")
                FormatarBotao(btnJaneiro, objListBalanco(e.Row.RowIndex).JaneiroAuditado)

                Dim btnFevereiro As Button = e.Row.FindControl("btnFevereiro")
                FormatarBotao(btnFevereiro, objListBalanco(e.Row.RowIndex).FevereiroAuditado)

                Dim btnMarco As Button = e.Row.FindControl("btnMarco")
                FormatarBotao(btnMarco, objListBalanco(e.Row.RowIndex).MarcoAuditado)

                Dim btnAbril As Button = e.Row.FindControl("btnAbril")
                FormatarBotao(btnAbril, objListBalanco(e.Row.RowIndex).AbrilAuditado)

                Dim btnMaio As Button = e.Row.FindControl("btnMaio")
                FormatarBotao(btnMaio, objListBalanco(e.Row.RowIndex).MaioAuditado)

                Dim btnJunho As Button = e.Row.FindControl("btnJunho")
                FormatarBotao(btnJunho, objListBalanco(e.Row.RowIndex).JunhoAuditado)

                Dim btnJulho As Button = e.Row.FindControl("btnJulho")
                FormatarBotao(btnJulho, objListBalanco(e.Row.RowIndex).JulhoAuditado)

                Dim btnAgosto As Button = e.Row.FindControl("btnAgosto")
                FormatarBotao(btnAgosto, objListBalanco(e.Row.RowIndex).AgostoAuditado)

                Dim btnSetembro As Button = e.Row.FindControl("btnSetembro")
                FormatarBotao(btnSetembro, objListBalanco(e.Row.RowIndex).SetembroAuditado)

                Dim btnOutubro As Button = e.Row.FindControl("btnOutubro")
                FormatarBotao(btnOutubro, objListBalanco(e.Row.RowIndex).OutrubroAuditado)

                Dim btnNovembro As Button = e.Row.FindControl("btnNovembro")
                FormatarBotao(btnNovembro, objListBalanco(e.Row.RowIndex).NovembroAuditado)

                Dim btnDezembro As Button = e.Row.FindControl("btnDezembro")
                FormatarBotao(btnDezembro, objListBalanco(e.Row.RowIndex).DezembroAuditado)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub FormatarBotao(btn As Button, Auditado As Boolean)
        If Auditado Then
            btn.BackColor = Drawing.Color.Red
            btn.BorderColor = Drawing.Color.Red
            btn.ForeColor = Drawing.Color.Yellow
            btn.Text = "Auditado"
        Else
            btn.BackColor = Drawing.Color.Green
            btn.ForeColor = Drawing.Color.White
            btn.BorderColor = Drawing.Color.Green
            btn.Text = "Aberto"
        End If
    End Sub


    Private Sub CarregarMovimento()
        Session.Remove("objListBalanco" & Hid.Value)
        Hid.Value = Guid.NewGuid().ToString
        objListBalanco = New ListBalancoAuditadoMes()
        objListBalanco.Carregar(ddlAno.SelectedValue)
        SessaoSalvaBalanco()

        gridMeses.DataSource = objListBalanco
        gridMeses.DataBind()
    End Sub

    Protected Sub ddlAno_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlAno.SelectedIndexChanged
        CarregarMovimento()
    End Sub

    Protected Sub Btn_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)

        Dim Mes As Integer
        Select Case btn.ID.ToUpper
            Case "BTNJANEIRO" : Mes = 1
            Case "BTNFEVEREIRO" : Mes = 2
            Case "BTNMARCO" : Mes = 3
            Case "BTNABRIL" : Mes = 4
            Case "BTNMAIO" : Mes = 5
            Case "BTNJUNHO" : Mes = 6
            Case "BTNJULHO" : Mes = 7
            Case "BTNAGOSTO" : Mes = 8
            Case "BTNSETEMBRO" : Mes = 9
            Case "BTNOUTUBRO" : Mes = 10
            Case "BTNNOVEMBRO" : Mes = 11
            Case "BTNDEZEMBRO" : Mes = 12
        End Select

        SessaoRecuperaBalanco()
        objListBalanco(row.RowIndex).Salvar(Mes)
        SessaoSalvaBalanco()

        gridMeses.DataSource = objListBalanco
        gridMeses.DataBind()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "BalancoAuditadoMes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class