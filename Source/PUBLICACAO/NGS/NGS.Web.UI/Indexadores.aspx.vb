Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Indexadores
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Indexadores", "ACESSAR") Then
                    CarregarIndexadores()
                    ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "", True)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Indexadores", "GRAVAR") Then
                If validacampos() Then
                    Dim sql As String = "INSERT Into Indexadores(Indexador_id, Descricao, Moeda) " & vbCrLf & _
                                        "                 Values('" & txtCodigo.Text.Trim & "'," & vbCrLf & _
                                        "                        '" & UCase(txtDescricao.Text.Trim) & "'," & vbCrLf & _
                                        "                        " & ddlMoeda.SelectedValue & ")" & vbCrLf

                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarIndexadores()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Indexadores", "ALTERAR") Then
                If validacampos() Then
                    Dim sql As String = "UPDATE Indexadores" & vbCrLf & _
                                         "  Set Descricao    = '" & txtDescricao.Text.Trim & "', " & vbCrLf & _
                                         "      Moeda        = '" & ddlMoeda.SelectedValue & "' " & vbCrLf & _
                                         "WHERE Indexador_Id = '" & txtCodigo.Text.Trim & "' " & vbCrLf

                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarIndexadores()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para atualizar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Indexadores", "EXCLUIR") Then
                If GridIndexadores.SelectedIndex = -1 OrElse String.IsNullOrWhiteSpace(txtCodigo.Text.Trim) Then
                    MsgBox(Me.Page, "Selecione o indexador que deseja excluir")
                Else
                    Dim sql As String = "DELETE FROM Indexadores" & vbCrLf & _
                                          " WHERE Indexador_Id = '" & txtCodigo.Text.Trim & "' " & vbCrLf

                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarIndexadores()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridIndexadores_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim moeda As String = ""
            If Not String.IsNullOrWhiteSpace(Server.HtmlDecode(GridIndexadores.SelectedRow.Cells(2).Text())) Then
                moeda = Server.HtmlDecode(GridIndexadores.SelectedRow.Cells(2).Text()).Split("-")(0)
            End If

            txtCodigo.Text = Server.HtmlDecode(GridIndexadores.SelectedRow.Cells(1).Text())
            ddlMoeda.SelectedValue = moeda
            txtDescricao.Text = Server.HtmlDecode(GridIndexadores.SelectedRow.Cells(3).Text())
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
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
            If Funcoes.VerificaPermissao("Indexadores", "RELATORIO") Then
                Dim sql As String = "Select REPLICATE('0', 3 - LEN(CAST(Indexador_Id AS varchar))) + CAST(Indexador_Id AS varchar) as Codigo, Descricao From Indexadores Order by Indexador_Id"
                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Indexadores.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
               
                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Indexadores")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub CarregarIndexadores(Optional ByVal moeda As String = "")
        Dim sql As String = " Select i.Indexador_Id as Codigo, cast(m.Moeda_Id as varchar) + '-' + m.Descricao as Moeda, i.Descricao " & vbCrLf & _
                            "   from Indexadores i                                       " & vbCrLf & _
                            "   Left Join Moedas m                                       " & vbCrLf & _
                            "     on m.Moeda_Id = i.Moeda                                " & vbCrLf

        If Not String.IsNullOrWhiteSpace(moeda) Then
            sql &= " where m.Moeda_Id = " & moeda & vbCrLf
        End If

        GridIndexadores.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
        GridIndexadores.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = String.Empty
        txtDescricao.Text = String.Empty
        ddlMoeda.SelectedValue = String.Empty

        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function validacampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text.Trim) Then
            MsgBox(Me.Page, "Informe o código.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlMoeda.SelectedValue) Then
            MsgBox(Me.Page, "Informe a moeda.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text.Trim) Then
            MsgBox(Me.Page, "Informe a descrição.")
            Return False
        ElseIf txtDescricao.Text.Trim.Length < 3 Then
            MsgBox(Me.Page, "Informe a descrição, com no mínimo três caracteres.")
            Return False
        End If
        Return True
    End Function

End Class