Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RotinasDeImpressao
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RotinasDeImpressao", "ACESSAR") Then
                CarregarRotinasDeImpressao()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarRotinasDeImpressao()
        If Globais.GPermiteLeitura = "S" Then
            Sql = "  SELECT Rotina_Id as Rotina, Descricao " & _
                            " FROM RotinasDeImpressao " & _
                            " ORDER BY Rotina_id"

            GridRotinasDeImpressao.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridRotinasDeImpressao.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtRotina.Text = ""
        txtDescricao.Text = ""
        txtRotina.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridRotinasDeImpressao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtRotina.Text = GridRotinasDeImpressao.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridRotinasDeImpressao.SelectedRow.Cells(2).Text()
            txtRotina.Enabled = False
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
            If Funcoes.VerificaPermissao("RotinasDeImpressao", "GRAVAR") Then
                Sql = "INSERT Into RotinasDeImpressao(Rotina_Id, Descricao) " & vbCrLf & _
                    " Values('" & txtRotina.Text & "' " & vbCrLf & _
                    ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Dados Inseridos com Sucesso.", eTitulo.Sucess)
                Else
                    Limpar()
                    CarregarRotinasDeImpressao()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para gravar Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("RotinasDeImpressao", "ALTERAR") Then
                Sql = "UPDATE RotinasDeImpressao" & vbCrLf & _
                    " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                    " WHERE Rotina_Id = '" & txtRotina.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Dados Atualizados com Sucesso.", eTitulo.Sucess)
                Else
                    Limpar()
                    CarregarRotinasDeImpressao()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para alterar Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("RotinasDeImpressao", "EXCLUIR") Then
                Sql = "DELETE FROM RotinasDeImpressao" & vbCrLf & _
                   "        WHERE Rotina_Id = '" & txtRotina.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, "Informações Excluída com Sucesso.", eTitulo.Sucess)
                Else
                    Limpar()
                    CarregarRotinasDeImpressao()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RotinasDeImpressao", "RELATORIO") Then
                Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                Sql = "Select Rotina_ID as Codigo, Descricao From RotinasDeImpressao Order by Rotina_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Rotinas De Impressão")
                parameters.Add("Codigo", "Rotina")
                parameters.Add("Descricao", "Descrição")
                Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                parameters.Add("Empresa", objEmpresa.Nome.Trim())
                parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RotinasDeImpressao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class