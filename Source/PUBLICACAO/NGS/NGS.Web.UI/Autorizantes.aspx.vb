Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Autorizantes
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim ds As DataSet
    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Autorizantes", "ACESSAR") Then
                CarregarAutorizantes()
                CarregarSubstitutos()
                CarregarAutorizantesGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarAutorizantes()
        If Funcoes.VerificaPermissao("Autorizantes", "LEITURA") Then
            Sql = "Select Usuario_Id as Autorizante From Usuarios Order By Usuario_Id"

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
                DdlAutorizantes.Items.Add(New ListItem(UCase(Dr("Autorizante")), UCase(Dr("Autorizante"))))
            Next
            DdlAutorizantes.Items.Insert(0, "")
            DdlAutorizantes.SelectedIndex = 0
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub CarregarSubstitutos()
        If Funcoes.VerificaPermissao("Autorizantes", "LEITURA") Then
            Sql = "Select Usuario_Id as Substituto From Usuarios Order By Usuario_Id"

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
                DdlSubstitutos.Items.Add(New ListItem(UCase(Dr("Substituto")), UCase(Dr("Substituto"))))
            Next
            DdlSubstitutos.Items.Insert(0, "")
            DdlSubstitutos.SelectedIndex = 0
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub CarregarAutorizantesGrid()
        If Funcoes.VerificaPermissao("Autorizantes", "LEITURA") Then
            Sql = "   SELECT  UPPER(Autorizantes.Autorizante_Id) AS Autorizante, ISNULL(Autorizantes.Substituto, '') AS Substituto, Autorizantes.Nivel, " & vbCrLf & _
                  "           CASE Autorizantes.CotacaoDePreco WHEN 'S' THEN 'SIM' ELSE 'NÃO' END Cotacao " & vbCrLf & _
                  "     FROM  Autorizantes INNER JOIN " & vbCrLf & _
                  "           Usuarios ON Autorizantes.Autorizante_Id = Usuarios.Usuario_Id " & vbCrLf & _
                  " ORDER BY  Autorizantes.Autorizante_Id "

            GridAutorizantes.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridAutorizantes.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub Limpar()
        DdlAutorizantes.SelectedIndex = 0
        DdlSubstitutos.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            DdlAutorizantes.SelectedValue = GridAutorizantes.SelectedRow.Cells(1).Text()
            If RTrim(GridAutorizantes.SelectedRow.Cells(2).Text()) <> "&nbsp;" Then
                DdlSubstitutos.SelectedValue = RTrim(GridAutorizantes.SelectedRow.Cells(2).Text())
            End If

            DdlNivel.SelectedValue = GridAutorizantes.SelectedRow.Cells(3).Text()
            DdlCotacao.SelectedValue = GridAutorizantes.SelectedRow.Cells(4).Text()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarAutorizantesGrid()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Autorizantes", "GRAVAR") Then
                Sql = "INSERT Into Autorizantes(Autorizante_Id, Substituto, Nivel, CotacaoDePreco) " & vbCrLf & _
                      " Values('" & UCase(DdlAutorizantes.SelectedValue) & "'" & vbCrLf & _
                      ",'" & UCase(DdlSubstitutos.SelectedValue) & "'" & vbCrLf & _
                      "," & DdlNivel.SelectedValue & vbCrLf & _
                      ",'" & UCase(DdlCotacao.SelectedValue) & "')" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Erro ao gravar registro.")
                Else
                    Limpar()
                    CarregarAutorizantesGrid()
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
            If Funcoes.VerificaPermissao("Autorizantes", "ALTERAR") Then
                Sql = "UPDATE Autorizantes" & vbCrLf & _
                      " Set Substituto = '" & UCase(DdlSubstitutos.SelectedValue) & "' " & vbCrLf & _
                      ",    Nivel= " & DdlNivel.SelectedValue & vbCrLf & _
                      ",    CotacaoDePreco = '" & DdlCotacao.SelectedValue & "' " & vbCrLf & _
                      " WHERE Autorizante_Id = '" & DdlAutorizantes.SelectedValue & "'" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Erro ao atualizar o registro.")
                Else
                    Limpar()
                    CarregarAutorizantesGrid()
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
            If Funcoes.VerificaPermissao("Autorizantes", "EXCLUIR") Then
                Sql = "DELETE FROM Autorizantes" & vbCrLf & _
                      " WHERE Autorizante_Id = '" & UCase(DdlAutorizantes.SelectedValue) & "'" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Erro ao atualizar o registro.")
                Else
                    Limpar()
                    CarregarAutorizantesGrid()
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
            If Funcoes.VerificaPermissao("Autorizantes", "RELATORIO") Then
                Sql = "   SELECT  UPPER(Autorizantes.Autorizante_Id) AS Autorizante, ISNULL(Autorizantes.Substituto, '') AS Substituto, Autorizantes.Nivel, " & vbCrLf & _
                      "           CASE Autorizantes.CotacaoDePreco WHEN 'S' THEN 'SIM' ELSE 'NÃO' END Cotacao " & vbCrLf & _
                      "     FROM  Autorizantes INNER JOIN " & vbCrLf & _
                      "           Usuarios ON Autorizantes.Autorizante_Id = Usuarios.Usuario_Id " & vbCrLf & _
                      " ORDER BY  Autorizantes.Autorizante_Id " & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Autorizantes")

                Dim param As New Dictionary(Of String, Object)
                param.Add("ConsultaParametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_Autorizantes", eExportType.PDF, param)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(DdlAutorizantes.SelectedValue) Then
            param &= "Parametros:" & vbCrLf & "Autorizante: " & DdlAutorizantes.SelectedValue & " "
        End If
        If Not String.IsNullOrWhiteSpace(DdlSubstitutos.SelectedValue) Then
            param &= "Substituto: " & DdlSubstitutos.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlNivel.SelectedValue) Then
            param &= "Nível: " & DdlNivel.SelectedValue & " "
        End If
        If Not String.IsNullOrWhiteSpace(DdlCotacao.SelectedValue) Then
            param &= "Cotação: " & DdlCotacao.SelectedValue
        End If

        Return param
    End Function

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Try
                Funcoes.Ajuda(Me.Page, "Autorizantes")
            Catch ex As Exception
                MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class