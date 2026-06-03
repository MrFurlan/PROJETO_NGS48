Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SeriesFiscais
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("SeriesFiscais", "ACESSAR") Then
            CarregarSeriesFiscais()
            Limpar()
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
            '    Exit Sub
            'End If
        End If
    End Sub

    Private Sub CarregarSeriesFiscais()
        Sql = "  SELECT Codigo_id as Codigo, Descricao " & vbCrLf & _
              " FROM SeriesFiscais " & vbCrLf & _
              " ORDER BY Codigo_Id" & vbCrLf

        GridSeries.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridSeries.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridSeries_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridSeries.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridSeries.SelectedRow.Cells(2).Text()
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
            'If Funcoes.VerificaPermissao("SeriesFiscais", "GRAVAR") Then
            Sql = "INSERT Into SeriesFiscais (Codigo_Id, Descricao) " & vbCrLf & _
                  " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                  ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
            Else
                Limpar()
                CarregarSeriesFiscais()
            End If
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.")
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            'If Funcoes.VerificaPermissao("SeriesFiscais", "ALTERAR") Then
            Sql = "UPDATE SeriesFiscais" & vbCrLf & _
                  " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                  " WHERE Codigo_Id = '" & txtCodigo.Text & "' " & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
            Else
                Limpar()
                CarregarSeriesFiscais()
            End If
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            'If Funcoes.VerificaPermissao("SeriesFiscais", "EXCLUIR") Then
            Sql = "DELETE FROM SeriesFiscais"
            Sql &= " WHERE Codigo_Id = '" & txtCodigo.Text & "' "
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
            Else
                Limpar()
                CarregarSeriesFiscais()
            End If
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            'If Funcoes.VerificaPermissao("SeriesFiscais", "RELATORIO") Then

            Dim Titulo As String = "Séries Fiscais"
            Dim Codigo As String = "Código"
            Dim Descricao As String = "Descrição"

            Titulo = "Séries Fiscais"
            Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo, Descricao From SeriesFiscais Order by Codigo_Id"

            DS = Banco.ConsultaDataSet(Sql, "Tabelas")

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Titulo", Titulo)
            parameters.Add("Codigo", Codigo)
            parameters.Add("Descricao", Descricao)
            Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
            parameters.Add("Empresa", objEmpresa.Nome.Trim())
            parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

            Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SeriesFiscais")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class