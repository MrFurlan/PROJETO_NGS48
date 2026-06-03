Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CarteiraDoTitulo
    Inherits BasePage

    Dim objCarteiraDoTitulo As [Lib].Negocio.CarteiraDoTitulo
    Private objCarteiras As ListCarteiraDoTitulo

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CarteiraDoTitulo", "ACESSAR") Then
                    ddl.Carregar(ddlBanco, CarregarDDL.Tabela.Bancos, "", True)
                    AtualizarGrid()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Gestao.aspx")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridCarteiraDoTitulo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False

            SessaoRecuperaCarteiras()

            txtCodigo.Text = objCarteiras(gridCarteiraDoTitulo.SelectedIndex).Codigo
            txtDescricao.Text = objCarteiras(gridCarteiraDoTitulo.SelectedIndex).Descricao

            If objCarteiras(gridCarteiraDoTitulo.SelectedIndex).CodigoDoBanco <> 0 Then
                ddlBanco.SelectedIndex = objCarteiras(gridCarteiraDoTitulo.SelectedIndex).CodigoDoBanco
            Else
                ddlBanco.SelectedIndex = 0
            End If

            chkFluxoDeCaixa.Checked = objCarteiras(gridCarteiraDoTitulo.SelectedIndex).FluxoDeCaixa
            chkDuplicata.Checked = objCarteiras(gridCarteiraDoTitulo.SelectedIndex).EmiteDuplicata

            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDoTitulo", "GRAVAR") Then
                If Not IsNumeric(txtCodigo.Text) Then
                    MsgBox(Me.Page, "Informe Um codigo Valido Para Cateira.")
                    Exit Sub
                ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
                    MsgBox(Me.Page, "Irforme a descrição.")
                    Exit Sub
                End If

                SessaoRecuperaCarteiras()
                Dim i = objCarteiras.Find(Function(s) s.Codigo = txtCodigo.Text)

                If Not i Is Nothing Then
                    MsgBox(Me.Page, "Este codigo já esta em uso")
                    Exit Sub
                End If

                Gravar_Registro("I")


            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDoTitulo", "ALTERAR") Then
                Gravar_Registro("U")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDoTitulo", "EXCLUIR") Then
                Gravar_Registro("D")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
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
            If Funcoes.VerificaPermissao("CarteiraDoTitulo", "RELATORIO") Then
                SessaoRecuperaCarteiras()

                Dim ds As DataSet = getDataSet()
                Dim parameters As New Dictionary(Of String, Object)

                Funcoes.BindReport(Me.Page, ds, "Cr_CarteiraDoTitulo", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão emitir relatório!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Public Sub SessaoRecuperaCarteiras()
        objCarteiras = Session("objCarteiras")
    End Sub

    Public Sub SessaoSalvaCarteiras()
        Session("objCarteiras") = objCarteiras
    End Sub

    Private Sub AtualizarGrid()
        objCarteiras = New ListCarteiraDoTitulo()
        gridCarteiraDoTitulo.DataSource = objCarteiras.ToArray()
        gridCarteiraDoTitulo.DataBind()

        SessaoSalvaCarteiras()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        ddlBanco.SelectedIndex = 0
        txtCodigo.Enabled = True
        chkDuplicata.Checked = False
        chkFluxoDeCaixa.Checked = False
        txtCodigo.Focus()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        gridCarteiraDoTitulo.SelectedIndex = -1
    End Sub

    Private Sub Gravar_Registro(ByVal tipo As String)
        objCarteiraDoTitulo = New [Lib].Negocio.CarteiraDoTitulo()

        objCarteiraDoTitulo.Codigo = txtCodigo.Text
        objCarteiraDoTitulo.Descricao = txtDescricao.Text
        objCarteiraDoTitulo.CodigoDoBanco = IIf(ddlBanco.SelectedIndex > 0, ddlBanco.SelectedValue, 0)
        objCarteiraDoTitulo.FluxoDeCaixa = chkFluxoDeCaixa.Checked
        objCarteiraDoTitulo.EmiteDuplicata = chkDuplicata.Checked
        objCarteiraDoTitulo.IUD = tipo
        If objCarteiraDoTitulo.Salvar Then
            Limpar()
            AtualizarGrid()

            Select Case tipo
                Case "I"
                    MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
                Case "U"
                    MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                Case "D"
                    MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
            End Select
        Else
            Select Case tipo
                Case "I"
                    MsgBox(Me.Page, "Erro ao salvar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
                Case "U"
                    MsgBox(Me.Page, "Erro ao alterar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
                Case "D"
                    MsgBox(Me.Page, "Erro ao excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
            End Select
        End If
    End Sub

    Private Function getParameters() As String
        Dim text As String = ""

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            text &= "Código: " & txtCodigo.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            text &= "Descrição: " & txtDescricao.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlBanco.SelectedValue) Then
            text &= "Banco: " & ddlBanco.SelectedValue & vbCrLf
        End If

        Return text
    End Function

    Private Function getDataSet() As DataSet
        Dim sql = "SELECT	C.Carteira_Id [Codigo], C.Descricao, convert(varchar, C.Banco) + ' - ' + B.Descricao AS Banco,   " & vbCrLf & _
                "		isnull(C.FluxoDeCaixa,0) AS FluxoDeCaixa, isnull(C.EmiteDuplicata,0) as EmiteDuplicata           " & vbCrLf & _
                "    FROM Carteira AS C                                                                                  " & vbCrLf & _
                "         LEFT JOIN Bancos AS B                                                                          " & vbCrLf & _
                "                ON B.Banco_id = C.Banco                                                                 " & vbCrLf & _
                "   Where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "    And C.Carteira_Id = " & txtCodigo.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "    And c.Descricao like '" & txtDescricao.Text & "%'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlBanco.SelectedValue) Then
            sql &= "    And isnull(C.Banco,0) = " & ddlBanco.SelectedValue & vbCrLf
        End If

        sql &= "   ORDER BY C.Carteira_Id" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "Carteira")
    End Function

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CarteiraDoTitulo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class
