Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ClasseToxicologica
    Inherits BasePage

    Dim objClasseToxicologica As [Lib].Negocio.ClasseToxicologica

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ClasseToxicologica", "ACESSAR") Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridClasseToxicologica_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False
            txtCodigo.Text = gridClasseToxicologica.SelectedRow.Cells(1).Text()
            txtDescricao.Text = gridClasseToxicologica.SelectedRow.Cells(2).Text()
            txtNomeCor.Text = gridClasseToxicologica.SelectedRow.Cells(3).Text()
            txtCor.Text = gridClasseToxicologica.SelectedRow.Cells(4).Text()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridClasseToxicologica_PreRender(sender As Object, e As EventArgs)
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If gridClasseToxicologica.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            gridClasseToxicologica.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            gridClasseToxicologica.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If

    End Sub

    Private Sub AtualizarGrid()
        If Funcoes.VerificaPermissao("ClasseToxicologica", "LEITURA") Then
            Dim Lista As New [Lib].Negocio.ListClasseToxicologica(True)
            gridClasseToxicologica.DataSource = Lista.ToArray
            gridClasseToxicologica.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub Limpar()
        objClasseToxicologica = New [Lib].Negocio.ClasseToxicologica
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtNomeCor.Text = ""
        txtCor.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtCodigo.Enabled = True

        Dim textAnotacion As TextBox = Me.Page.FindControl(txtCodigo.UniqueID)
        Dim scriptManager As ScriptManager = scriptManager.GetCurrent(Me.Page)
        textAnotacion.Focus()
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql = "Select ClasseTox_Id AS Codigo, Descricao, isnull(CorClasse,'') as CorClasse, isnull(CodCorClasse,'') as CodCorClasse  From ClasseToxicologica" & vbCrLf & _
                "   Where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And ClasseTox_Id = " & txtCodigo.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao like '" & txtDescricao.Text & "%'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCor.Text) Then
            sql &= "And CodCorClasse like '" & txtCor.Text & "%'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNomeCor.Text) Then
            sql &= "And CorClasse like '" & txtNomeCor.Text & "%'" & vbCrLf
        End If

        sql &= "   ORDER BY ClasseTox_Id" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "ClasseToxicologica")
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ClasseToxicologica", "GRAVAR") Then
                objClasseToxicologica = New [Lib].Negocio.ClasseToxicologica

                objClasseToxicologica.Codigo = txtCodigo.Text
                objClasseToxicologica.Descricao = txtDescricao.Text
                objClasseToxicologica.CodigoCor = txtCor.Text
                objClasseToxicologica.Cor = txtNomeCor.Text
                objClasseToxicologica.IUD = "I"
                If objClasseToxicologica.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ClasseToxicologica", "ALTERAR") Then
                objClasseToxicologica = New [Lib].Negocio.ClasseToxicologica

                objClasseToxicologica.Codigo = gridClasseToxicologica.SelectedRow.Cells(1).Text()
                objClasseToxicologica.Descricao = txtDescricao.Text
                objClasseToxicologica.CodigoCor = txtCor.Text
                objClasseToxicologica.Cor = txtNomeCor.Text
                objClasseToxicologica.IUD = "U"
                If objClasseToxicologica.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ClasseToxicologica", "EXCLUIR") Then
                objClasseToxicologica = New [Lib].Negocio.ClasseToxicologica

                objClasseToxicologica.Codigo = gridClasseToxicologica.SelectedRow.Cells(1).Text()
                objClasseToxicologica.IUD = "D"
                If objClasseToxicologica.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ClasseToxicologica", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()

                Funcoes.BindReport(Me.Page, ds, "Cr_ClasseToxicologica", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ClasseToxicologica")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class