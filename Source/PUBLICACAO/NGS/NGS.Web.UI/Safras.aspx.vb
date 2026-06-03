Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Safras
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Safras", "ACESSAR") Then
                    CarregarSafras()
                    CarregarGrupo()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridSafras_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtSafra.Text = Server.HtmlDecode(GridSafras.SelectedRow.Cells(1).Text())
            DdlGrupo.SelectedValue = Server.HtmlDecode(GridSafras.SelectedRow.Cells(2).Text())
            CarregarProduto()
            DdlProduto.SelectedValue = Server.HtmlDecode(GridSafras.SelectedRow.Cells(3).Text())

            txtInicioSafra.Text = Server.HtmlDecode(GridSafras.SelectedRow.Cells(5).Text())
            txtVencimento.Text = Server.HtmlDecode(GridSafras.SelectedRow.Cells(5).Text())
            txtTaxa.Text = Server.HtmlDecode(GridSafras.SelectedRow.Cells(6).Text())
            txtObservacao.Text = Server.HtmlDecode(GridSafras.SelectedRow.Cells(7).Text())
            txtSafra.Enabled = False
            txtObservacao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Safras", "GRAVAR") Then
                If validacampos() Then
                    Sql = "INSERT Into Safras (Safra_id, Produto,InicioDeSafra, Vencimento, Taxa, Observacao)   " & vbCrLf & _
                                          "             Values('" & UCase(txtSafra.Text) & "' " & vbCrLf & _
                                          ", '" & DdlProduto.SelectedValue & "'" & vbCrLf & _
                                          ", '" & txtInicioSafra.Text.Trim().ToSqlDate() & "'" & vbCrLf & _
                                          ", '" & txtVencimento.Text.ToSqlDate() & "'" & vbCrLf & _
                                          ", " & Replace(txtTaxa.Text, ",", ".") & vbCrLf & _
                                          ", '" & txtObservacao.Text.ToUpper() & "')" & vbCrLf

                    If Banco.GravaBanco(Sql) Then
                        MsgBox(Me.Page, "Registro incluso com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarSafras()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Safras", "ALTERAR") Then
                If validacampos() Then
                    Sql = "UPDATE Safras" & vbCrLf & _
                          " Set Produto    ='" & DdlProduto.SelectedValue & "'" & vbCrLf & _
                          ", InicioDeSafra ='" & txtInicioSafra.Text.ToSqlDate() & "' " & vbCrLf & _
                          ", Vencimento    ='" & txtVencimento.Text.ToSqlDate() & "' " & vbCrLf & _
                          ", Taxa          = " & Replace(txtTaxa.Text, ",", ".") & vbCrLf & _
                          ", Observacao    ='" & txtObservacao.Text.ToUpper() & "'" & vbCrLf & _
                          " WHERE Safra_Id ='" & txtSafra.Text & "' " & vbCrLf

                    If Banco.GravaBanco(Sql) Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarSafras()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Safras", "EXCLUIR") Then
                If String.IsNullOrWhiteSpace(txtSafra.Text) Then
                    MsgBox(Me.Page, "Informe a safra.")
                Else
                    Sql = "DELETE FROM Safras" & vbCrLf & _
                          " WHERE Safra_Id = '" & txtSafra.Text & "' "

                    If Banco.GravaBanco(Sql) Then
                        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarSafras()
                    End If
                End If
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
            If Funcoes.VerificaPermissao("Safras", "RELATORIO") Then

                Sql = "Select Safra_Id as Codigo, Observacao as Descricao From Safras Order by Safra_Id"
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Safras")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar o relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Safras")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CarregarSafras()
        Sql = "SELECT Safras.Safra_Id as Safra, Produtos.Grupo, Safras.Produto, " & vbCrLf & _
              "       isnull(Safras.InicioDeSafra,Safras.Vencimento) as InicioDeSafra, Safras.Vencimento," & vbCrLf & _
              "       Safras.Taxa, Safras.Observacao" & vbCrLf & _
              "  FROM Safras " & vbCrLf & _
              " INNER JOIN Produtos " & vbCrLf & _
              "    ON Safras.Produto = Produtos.Produto_Id" & vbCrLf & _
              " ORDER BY Safra_id" & vbCrLf

        GridSafras.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridSafras.DataBind()
    End Sub

    Private Sub CarregarGrupo()
        DdlGrupo.Items.Clear()

        Sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao" & _
              " FROM         GruposDeEstoques INNER JOIN" & _
              "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & _
              " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) order by GruposDeEstoques.Descricao"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlGrupo.Items.Add(New ListItem(Dr("Descricao"), Dr("Grupo_Id")))
        Next

        DdlGrupo.Items.Insert(0, "")
        DdlGrupo.SelectedIndex = 0
    End Sub

    Private Sub CarregarProduto()
        DdlProduto.Items.Clear()

        Sql = "SELECT Produto_Id, Nome FROM Produtos WHERE Grupo = '" & DdlGrupo.SelectedValue & "'" & _
              " Order by Nome"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlProduto.Items.Add(New ListItem(Dr("Nome"), Dr("Produto_Id")))
        Next

        DdlProduto.Items.Insert(0, "")
        DdlProduto.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        txtSafra.Text = String.Empty
        DdlGrupo.SelectedValue = String.Empty
        DdlProduto.SelectedValue = String.Empty
        txtInicioSafra.Text = Date.Today.ToString("dd/MM/yyyy")
        txtVencimento.Text = Date.Today.ToString("dd/MM/yyyy")
        txtTaxa.Text = String.Empty
        txtObservacao.Text = String.Empty
        txtSafra.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function validacampos() As Boolean
        If String.IsNullOrWhiteSpace(txtSafra.Text) Then
            MsgBox(Me.Page, "Informe a safra.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue()) Then
            MsgBox(Me.Page, "Informe o grupo.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlProduto.SelectedValue()) Then
            MsgBox(Me.Page, "Informe o produto.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtInicioSafra.Text) Then
            MsgBox(Me.Page, "Informe o inicio da safra.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtVencimento.Text) Then
            MsgBox(Me.Page, "Informe o vencimento da safra.")
            Return False
        ElseIf Not IsDate(txtInicioSafra.Text) Then
            MsgBox(Me.Page, "Inicio da safra, não é uma data válida.")
            Return False
        ElseIf Not IsDate(txtVencimento.Text) Then
            MsgBox(Me.Page, "Vencimento da safra não é uma data válida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtTaxa.Text) Then
            MsgBox(Me.Page, "Informe a taxa.")
            Return False
        End If
        Return True
    End Function

#End Region

End Class