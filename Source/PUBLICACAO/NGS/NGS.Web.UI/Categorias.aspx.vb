Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Categorias
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim Sql As String
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Categoria", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                CarregarCategorias()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarCategorias()
        Sql = "  SELECT Categoria_Id as Codigo, Descricao, Funrural, BasePisCofins" & _
                                   " FROM Categorias " & _
                                   " ORDER BY Categoria_id"
        GridCategorias.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridCategorias.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtFunrural.Text = ""
        txtBasePisCofins.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            param &= "Código: " & txtCodigo.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtFunrural.Text) Then
            param &= "Funrural: " & txtFunrural.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtBasePisCofins.Text) Then
            param &= "Base Pis Cofins: " & txtBasePisCofins.Text
        End If

        Return param
    End Function

    Protected Sub GridCategorias_PreRender(sender As Object, e As EventArgs)
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If GridCategorias.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            GridCategorias.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            GridCategorias.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If

    End Sub

    Protected Sub GridCategorias_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridCategorias.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridCategorias.SelectedRow.Cells(2).Text()
            txtFunrural.Text = GridCategorias.SelectedRow.Cells(3).Text()
            txtBasePisCofins.Text = GridCategorias.SelectedRow.Cells(4).Text()
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Categoria", "GRAVAR") Then
                Sql = "INSERT Into Categorias (Categoria_Id, Descricao, Funrural, BasePisCofins) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "' " & vbCrLf & _
                      ", " & Replace(txtFunrural.Text, ",", ".") & vbCrLf & _
                      "," & Replace(txtBasePisCofins.Text, ",", ".") & ")" & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarCategorias()
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
            If Funcoes.VerificaPermissao("Categoria", "ALTERAR") Then
                Sql = "UPDATE Categorias" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      ", Funrural = " & Replace(txtFunrural.Text, ",", ".") & vbCrLf & _
                      ", BasePisCofins = " & Replace(txtBasePisCofins.Text, ",", ".") & vbCrLf & _
                      " WHERE Categoria_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarCategorias()
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
            If Funcoes.VerificaPermissao("Categoria", "EXCLUIR") Then
                Sql = "DELETE FROM Categorias" & vbCrLf & _
                      " WHERE Categoria_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarCategorias()
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
            If Funcoes.VerificaPermissao("Categoria", "RELATORIO") Then
                Sql = "  SELECT Categoria_Id as Codigo, Descricao, Funrural, BasePisCofins as Cofins" & vbCrLf & _
                      " FROM Categorias " & vbCrLf & _
                      " ORDER BY Categoria_Id" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Categorias")

                Funcoes.BindReport(Me.Page, DS, "Cr_Categorias", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Categoria")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class