Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TabelaDeComissao
    Inherits BasePage

    Dim objTabelaDeComissao As [Lib].Negocio.TabelaDeComissao
    Dim objFaixaDeComissao As [Lib].Negocio.FaixaDeComissao

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TabelaDeComissao", "ACESSAR") Then
                If Funcoes.VerificaPermissao("TabelaDeComissao", "LEITURA") Then
                    CargaSafra()
                    Limpar()
                    AtualizarGrid()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        Else
            If Not Request("__EVENTARGUMENT") Is Nothing Then
                If Request("__EVENTARGUMENT") = "AlterarTxC" Then IniciarTxC("U")
                If Request("__EVENTARGUMENT") = "ExcluirTxC" Then IniciarTxC("D")
                If Request("__EVENTARGUMENT") = "AlterarFxC" Then IniciarFxC("U")
                If Request("__EVENTARGUMENT") = "ExcluirFxC" Then IniciarFxC("D")
            End If
        End If
    End Sub

    Protected Sub gridTabelaDeComissao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkNovoFaixa.Parent.Visible = True
            Dim tComissao = New [Lib].Negocio.TabelaDeComissao(gridTabelaDeComissao.SelectedRow.Cells(1).Text())
            txtCodigoTabela.Value = tComissao.Codigo
            txtCodigo.Text = tComissao.Codigo
            txtDescricao.Text = tComissao.Descricao
            ddlSafra.SelectedValue = tComissao.CodigoSafra
            lblTabela.Text = tComissao.Codigo & " - " & tComissao.Descricao
            AtualizarGridFaixa(tComissao)
            LimparCamposFaixa()
            lnkNovoFaixa.Parent.Visible = True
            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaSafra()
        Dim Safras As New [Lib].Negocio.ListSafra(True)
        ddlSafra.Items.Clear()
        ddlSafra.DataTextField = "Codigo"
        ddlSafra.DataValueField = "Codigo"
        ddlSafra.DataSource = Safras.ToArray()
        ddlSafra.DataBind()
    End Sub

    Private Sub IniciarTxC(ByVal Tipo As String)
        If txtCodigo.Text.Length = 0 Then
            MsgBox(Me.Page, "Código não foi Infomado")
        ElseIf txtCodigo.Text.Length > 0 AndAlso CInt(txtCodigo.Text) = 0 Then
            MsgBox(Me.Page, "Código deve ser maior que Zero")
        ElseIf txtDescricao.Text.Length = 0 Then
            MsgBox(Me.Page, "Desscrição não foi Infomada")
        Else
            If Tipo = "I" And Funcoes.VerificaPermissao("TabelaDeComissao", "GRAVAR") Or _
               Tipo = "U" And Funcoes.VerificaPermissao("TabelaDeComissao", "ALTERAR") Or _
               Tipo = "D" And Funcoes.VerificaPermissao("TabelaDeComissao", "EXCLUIR") Then
                objTabelaDeComissao = New [Lib].Negocio.TabelaDeComissao

                objTabelaDeComissao.Codigo = txtCodigo.Text
                objTabelaDeComissao.Descricao = txtDescricao.Text
                objTabelaDeComissao.CodigoSafra = ddlSafra.SelectedValue
                objTabelaDeComissao.IUD = Tipo

                Dim Lista As New [Lib].Negocio.ListFaixaDeComissao(objTabelaDeComissao)

                If objTabelaDeComissao.IUD = "D" AndAlso Lista.Count > 0 Then
                    MsgBox(Me.Page, "Tabela com Faixa de Comissão Cadastrada não pode ser Excluída.")
                ElseIf objTabelaDeComissao.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Erro ao Incluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "U"
                            MsgBox(Me.Page, "Erro ao Alterar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "D"
                            MsgBox(Me.Page, "Erro ao Excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End Select
                End If
            Else
                Select Case Tipo
                    Case "I"
                        MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
                    Case "U"
                        MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
                    Case "D"
                        MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
                End Select
            End If
        End If
    End Sub

    Private Sub IniciarFxC(ByVal Tipo As String)
        If txtFaixaInicial.Text.Length = 0 Then
            MsgBox(Me.Page, "Faixa inicial não foi Infomada")
        ElseIf txtFaixaFinal.Text.Length = 0 Then
            MsgBox(Me.Page, "Faixa final não foi Infomada")
        ElseIf txtIndice.Text.Length = 0 Then
            MsgBox(Me.Page, "Indice não foi Infomado")
        ElseIf txtFaixaInicial.Text.Length > 0 AndAlso CDec(txtFaixaInicial.Text) = 0 Then
            MsgBox(Me.Page, "Faixa inicial deve ser maior que Zero")
        ElseIf txtFaixaFinal.Text.Length > 0 AndAlso CDec(txtFaixaFinal.Text) = 0 Then
            MsgBox(Me.Page, "Faixa final deve ser maior que Zero")
        ElseIf txtIndice.Text.Length > 0 AndAlso CDec(txtIndice.Text) = 0 Then
            MsgBox(Me.Page, "Indice deve ser maior que Zero")
        Else
            If Tipo = "I" And Funcoes.VerificaPermissao("TabelaDeComissao", "GRAVAR") Or _
               Tipo = "U" And Funcoes.VerificaPermissao("TabelaDeComissao", "ALTERAR") Or _
               Tipo = "D" And Funcoes.VerificaPermissao("TabelaDeComissao", "EXCLUIR") Then
                Dim tComissao As New [Lib].Negocio.TabelaDeComissao(txtCodigoTabela.Value)
                objFaixaDeComissao = New [Lib].Negocio.FaixaDeComissao(tComissao)
                objFaixaDeComissao.FaixaInicial = txtFaixaInicial.Text
                objFaixaDeComissao.FaixaFinal = txtFaixaFinal.Text
                objFaixaDeComissao.Indice = txtIndice.Text
                objFaixaDeComissao.Observacoes = txtObservacao.Text

                objFaixaDeComissao.IUD = Tipo

                If objFaixaDeComissao.IUD = "I" AndAlso objFaixaDeComissao.ConsultarFaixaInicial() Then
                    MsgBox(Me.Page, "Faixa Inicial já está Cadastrada.")
                ElseIf objFaixaDeComissao.Salvar Then
                    LimparCamposFaixa()
                    AtualizarGridFaixa(tComissao)
                    lnkNovoFaixa.Parent.Visible = True
                Else
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Erro ao Incluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "U"
                            MsgBox(Me.Page, "Erro ao Alterar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "D"
                            MsgBox(Me.Page, "Erro ao Excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End Select
                End If
            Else
                Select Case Tipo
                    Case "I"
                        MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
                    Case "U"
                        MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
                    Case "D"
                        MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
                End Select
            End If
        End If
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListTabelaDeComissao(True, "", "Codigo")
        gridTabelaDeComissao.DataSource = Lista.ToArray()
        gridTabelaDeComissao.DataBind()
    End Sub

    Private Sub Limpar()
        objTabelaDeComissao = New [Lib].Negocio.TabelaDeComissao
        txtCodigoTabela.Value = ""
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        ddlSafra.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        LimparCamposFaixa()
        lnkNovoFaixa.Parent.Visible = False
        lblTabela.Text = ""
        gridFaixaDeComissao.DataBind()
        TabContainer1.ActiveTabIndex = 0
        txtDescricao.Focus()
    End Sub

    Private Sub AtualizarGridFaixa(ByVal Tabela As [Lib].Negocio.TabelaDeComissao)
        Dim Lista As New [Lib].Negocio.ListFaixaDeComissao(Tabela)
        gridFaixaDeComissao.DataSource = Lista.ToArray()
        gridFaixaDeComissao.DataBind()
    End Sub

    Private Sub LimparCamposFaixa()
        txtFaixaInicial.Text = ""
        txtFaixaFinal.Text = ""
        txtIndice.Text = ""
        txtObservacao.Text = ""
        lnkAlterarFaixa.Parent.Visible = False
        lnkExcluirFaixa.Parent.Visible = True
        txtFaixaInicial.Enabled = True
    End Sub

    Protected Sub gridFaixaDeComissao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtFaixaInicial.Text = gridFaixaDeComissao.SelectedRow.Cells(1).Text()
            txtFaixaFinal.Text = gridFaixaDeComissao.SelectedRow.Cells(2).Text()
            txtIndice.Text = gridFaixaDeComissao.SelectedRow.Cells(3).Text()
            txtObservacao.Text = gridFaixaDeComissao.SelectedRow.Cells(4).Text()
            lnkNovoFaixa.Parent.Visible = False
            lnkAlterarFaixa.Parent.Visible = True
            lnkExcluirFaixa.Parent.Visible = True
            txtFaixaInicial.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            IniciarTxC("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            IniciarTxC("U")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("TabelaDeComissao", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Tabela De Comissão")
                parameters.Add("ParametrosConsulta", getParameters())

                Funcoes.BindReport(Me.Page, ds, "Cr_TabelaDeComissao", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParameters() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            param &= "Código: " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text
        End If
        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            param &= "And Safra: " & ddlSafra.SelectedValue & ""
        End If

        Return param
    End Function

    Private Function getDataSet() As DataSet
        Dim sql As String = "SELECT Codigo_Id As Codigo, Descricao, Safra" & vbCrLf & _
                            "  FROM TabelaDeComissao where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And Codigo_Id = " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao = '" & txtDescricao.Text & "'"
        End If
        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            sql &= "And Safra = '" & ddlSafra.SelectedValue & "'"
        End If

        sql &= "order by Codigo_Id"

        Return Banco.ConsultaDataSet(sql, "TabelaDeComissao")
    End Function

    Protected Sub lnkNovoFaixa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoFaixa.Click
        Try
            IniciarFxC("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirFaixa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirFaixa.Click
        Try
            IniciarFxC("D")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAlterarFaixa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAlterarFaixa.Click
        Try
            IniciarFxC("U")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparFaixa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparFaixa.Click
        Try
            LimparCamposFaixa()
            lnkLimparFaixa.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorioFaixa_Click(sender As Object, e As EventArgs) Handles lnkRelatorioFaixa.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaFaixa_Click(sender As Object, e As EventArgs) Handles lnkAjudaFaixa.Click
        Try
            Funcoes.Ajuda(Me.Page, "TabelaDeComissao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TabelaDeComissao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class