Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Impressoras
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Impressoras", "ACESSAR") Then
                pageLoad()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub grdImpressoras_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ds As DataSet
        Dim pCodigo As String = grdImpressoras.SelectedRow.Cells(1).Text()
        Try
            If Not String.IsNullOrWhiteSpace(pCodigo) Then
                Dim sql As String = "Select Rotina_Id, Descricao From RotinasDeImpressao where Rotina_Id = '" & pCodigo & "'"
                ds = Banco.ConsultaDataSet(sql, "RotinasDeImpressao")
                If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                    txtRotina.Text = ds.Tables(0).Rows(0)("Rotina_Id")
                    txtDescricao.Text = ds.Tables(0).Rows(0)("Descricao")
                End If
            End If
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub pageLoad()
        Dim sql As String = "Select Rotina_Id, Descricao From RotinasDeImpressao order by Rotina_Id"
        grdImpressoras.DataSource = Banco.ConsultaDataSet(sql, "RotinasDeImpressao")
        grdImpressoras.DataBind()
        Limpar()
    End Sub

    Private Function Validar() As Boolean
        Dim aux As Boolean = True
        If (String.IsNullOrWhiteSpace(txtRotina.Text)) Then
            aux = False
        End If
        If (String.IsNullOrWhiteSpace(txtDescricao.Text)) Then
            aux = False
        End If
        Return aux
    End Function

    Private Sub Incluir(ByVal Codigo As String, ByVal Descricao As String)
        Dim SqlArray As New ArrayList
        Dim sql As String = "INSERT INTO RotinasDeImpressao (Rotina_Id, Descricao) Values ('" & Codigo & "','" & Descricao & "')"
        SqlArray.Add(sql)
        If Not Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
        End If
    End Sub

    Private Sub Alterar(ByVal Codigo As String, ByVal Descricao As String)
        Dim SqlArray As New ArrayList
        Dim sql As String = "Update RotinasDeImpressao Set Descricao = '" & Descricao & "' Where Rotina_Id = '" & Codigo & "'"
        SqlArray.Add(sql)
        If Not Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Private Sub Excluir(ByVal Codigo As String)
        Dim SqlArray As New ArrayList
        Dim sql As String = "Delete From RotinasDeImpressao Where (Rotina_Id = '" & Codigo & "')"
        SqlArray.Add(sql)
        If Not Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Private Sub Relatorio()
        Try
            If Funcoes.VerificaPermissao("Impressoras", "RELATORIO") Then
                Dim ds As New DataSet
                Dim sql As String = "Select Rotina_Id as Codigo, Descricao From RotinasDeImpressao Order by Rotina_Id"
                ds = Banco.ConsultaDataSet(sql, "Tabelas")

                Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Impressoras")
                parameters.Add("Codigo", "Rotina")
                parameters.Add("Descricao", "Descrição")
                Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                parameters.Add("Empresa", objEmpresa.Nome.Trim())
                parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtRotina.Text = String.Empty
        txtDescricao.Text = String.Empty
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Impressoras", "GRAVAR") Then
                If Not Validar() Then
                    MsgBox(Me.Page, "É necessário preencher os campos obrigatórios.")
                    Exit Sub
                End If
                Incluir(txtRotina.Text.Trim(), txtDescricao.Text.Trim())
                pageLoad()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Not Validar() Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios.")
                Exit Sub
            End If
            Alterar(txtRotina.Text.Trim(), txtDescricao.Text.Trim())
            pageLoad()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Impressoras", "EXCLUIR") Then
                Excluir(txtRotina.Text.Trim())
                pageLoad()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir.")
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
            Relatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Impressoras")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class