Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class FormulacaoDoFito
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objFormulacaoFito As [Lib].Negocio.FormulacaoFito

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("FormulacaoDoFito", "ACESSAR") Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridFormulacaoFito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtCodigo.Enabled = False
            txtCodigo.Text = gridFormulacaoFito.SelectedRow.Cells(1).Text()
            txtDescricao.Text = gridFormulacaoFito.SelectedRow.Cells(2).Text()
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListFormulacaoFito(True)
        gridFormulacaoFito.DataSource = Lista.ToArray()
        gridFormulacaoFito.DataBind()
    End Sub

    Private Sub Limpar()
        objFormulacaoFito = New [Lib].Negocio.FormulacaoFito
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        txtCodigo.Focus()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("FormulacaoDoFito", "GRAVAR") Then
                objFormulacaoFito = New [Lib].Negocio.FormulacaoFito

                objFormulacaoFito.Codigo = txtCodigo.Text
                objFormulacaoFito.Descricao = txtDescricao.Text
                objFormulacaoFito.IUD = "I"
                If objFormulacaoFito.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("FormulacaoDoFito", "ALTERAR") Then
                objFormulacaoFito = New [Lib].Negocio.FormulacaoFito

                objFormulacaoFito.Codigo = gridFormulacaoFito.SelectedRow.Cells(1).Text()
                objFormulacaoFito.Descricao = txtDescricao.Text
                objFormulacaoFito.IUD = "U"
                If objFormulacaoFito.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("FormulacaoDoFito", "EXCLUIR") Then
                objFormulacaoFito = New [Lib].Negocio.FormulacaoFito

                objFormulacaoFito.Codigo = gridFormulacaoFito.SelectedRow.Cells(1).Text()
                objFormulacaoFito.IUD = "D"
                If objFormulacaoFito.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("FormulacaoDoFito", "RELATORIO") Then
                Sql = "Select FormulacaoFito_Id as Codigo, Descricao From FormulacaoFito order by FormulacaoFito_Id" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Formulação do Fito")
                parameters.Add("Codigo", "Código")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "FormulacaoDoFito")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class