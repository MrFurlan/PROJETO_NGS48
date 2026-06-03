Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Analises
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Analises", "ACESSAR") Then
                BindGridView()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub GrdAnalises_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim pCodigo As String = GrdAnalises.SelectedRow.Cells(1).Text()
        Selecionar(pCodigo)
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
    End Sub

    Private Sub Incluir(ByVal Analise As String, ByVal Descricao As String, ByVal IndiceMinimo As String, ByVal IndiceMaximo As String, ByVal Opcao As String)
        If Funcoes.VerificaPermissao("Analises", "GRAVAR") Then
            Dim SqlArray As New ArrayList
            Dim strSQL As String

            If IndiceMinimo = "" Then IndiceMinimo = "0"
            If IndiceMaximo = "" Then IndiceMaximo = "0"

            strSQL = "INSERT INTO Analises (Analise_Id, Descricao, IndiceMinimo, IndiceMaximo, Opcao)" & vbCrLf & _
                     "VALUES (" & Analise & ", '" & RTrim(Descricao) & "', " & Str(IndiceMinimo) & ", " & IndiceMaximo & "," & IIf(Opcao.Trim.Length = 0, "NULL", "'" & Opcao & "'") & ")"

            SqlArray.Add(strSQL)
            Banco.GravaBanco(SqlArray)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
        End If
    End Sub

    Private Function Alterar(ByVal Analise As String, ByVal Descricao As String, ByVal IndiceMinimo As String, ByVal IndiceMaximo As String, ByVal Opcao As String)
        If Funcoes.VerificaPermissao("Analises", "ALTERAR") Then
            Dim SqlArray As New ArrayList

            Dim strSQL As String
            If IndiceMinimo = "" Then IndiceMinimo = "0"
            If IndiceMaximo = "" Then IndiceMaximo = "0"

            strSQL = "UPDATE Analises SET " & vbCrLf & _
                     "   Descricao    ='" & Descricao.Trim & "'" & vbCrLf & _
                     "  ,IndiceMinimo = " & Str(IndiceMinimo) & vbCrLf & _
                     "  ,IndiceMaximo = " & Str(IndiceMaximo) & vbCrLf & _
                     "  ,Opcao        ='" & Opcao & "'" & vbCrLf & _
                     " WHERE Analise_Id = " & Analise

            SqlArray.Add(strSQL)
            Return Banco.GravaBanco(SqlArray)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para Alterar Registro")
        End If
        Return True
    End Function

    Private Function Excluir(ByVal Analise As String)
        If Funcoes.VerificaPermissao("Analises", "EXCLUIR") Then
            Dim SqlArray As New ArrayList

            Dim strSQL As String

            strSQL = "DELETE Analises" & vbCrLf & _
                     " WHERE Analise_Id = " & Analise

            SqlArray.Add(strSQL)
            Return Banco.GravaBanco(SqlArray)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para Excluir Registro")
        End If
        Return True
    End Function

    Private Sub Selecionar(ByVal pCodigo As String)
        Dim obj As New [Lib].Negocio.Analise(pCodigo)
        txtAnaliseId.Text = obj.Codigo
        txtDescricao.Text = obj.Descricao
        txtIndiceMinimo.Text = obj.IndiceMinimo
        txtIndiceMaximo.Text = obj.IndiceMaximo
        txtOpcao.Text = obj.Opcao
    End Sub

    Private Sub BindGridView()
        GrdAnalises.DataSource = New ListAnalise()
        GrdAnalises.DataBind()
    End Sub

    Private Sub Limpar()
        txtAnaliseId.Text = ""
        txtDescricao.Text = ""
        txtIndiceMinimo.Text = ""
        txtIndiceMaximo.Text = ""
        txtOpcao.Text = ""

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function Validar() As Boolean

        If (String.IsNullOrWhiteSpace(txtAnaliseId.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo analise.")
            Return False
        End If
        If (String.IsNullOrWhiteSpace(txtDescricao.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo descrição!")
            Return False
        End If
        If (String.IsNullOrWhiteSpace(txtIndiceMinimo.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo indíce mínimo!")
            Return False
        End If
        If (String.IsNullOrWhiteSpace(txtIndiceMaximo.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo indíce máximo!")
            Return False
        End If

        Return True
    End Function

    Private Function getDataSet() As DataSet
        Dim Sql As String
        Dim ds As New DataSet

        Sql = "SELECT Analise_Id, Descricao, IndiceMinimo, IndiceMaximo, Opcao " & vbCrLf & _
            "   FROM  Analises Order by Analise_Id"

        ds = Banco.ConsultaDataSet(Sql, "Analises")

        Return ds
    End Function

    Private Sub Relatorio()
        Try
            If Funcoes.VerificaPermissao("Analises", "RELATORIO") Then

                Dim Ds_Analises As DataSet = getDataSet()

                Funcoes.BindReport(Me.Page, Ds_Analises, "Cr_Analises", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Analises", "GRAVAR") Then
                If Validar() Then
                    Incluir(txtAnaliseId.Text, txtDescricao.Text, txtIndiceMinimo.Text, txtIndiceMaximo.Text, txtOpcao.Text)
                    BindGridView()
                    Limpar()
                    MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
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
            If Funcoes.VerificaPermissao("Analises", "ALTERAR") Then
                If Validar() Then
                    Alterar(txtAnaliseId.Text, txtDescricao.Text, txtIndiceMinimo.Text, txtIndiceMaximo.Text, txtOpcao.Text)
                    BindGridView()
                    Limpar()
                    MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
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
            If Funcoes.VerificaPermissao("Analises", "EXCLUIR") Then
                Excluir(txtAnaliseId.Text)
                BindGridView()
                Limpar()
                MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Relatorio()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Analises")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class