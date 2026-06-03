Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ClasseDeRisco
    Inherits BasePage

    Dim objClasseDeRisco As [Lib].Negocio.ClasseDeRisco

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ClasseDeRisco", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridClasseDeRisco_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False
            txtCodigo.Text = gridClasseDeRisco.SelectedRow.Cells(1).Text()
            txtDescricao.Text = gridClasseDeRisco.SelectedRow.Cells(2).Text()
            'ss(Page.FindControl("txtDescricao"))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        If Funcoes.VerificaPermissao("ClasseDeRisco", "LEITURA") Then
            Dim Lista As New [Lib].Negocio.ListClasseDeRisco(True)
            gridClasseDeRisco.DataSource = Lista.ToArray()
            gridClasseDeRisco.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub Limpar()
        objClasseDeRisco = New [Lib].Negocio.ClasseDeRisco
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtCodigo.Enabled = True
        Dim textAnotacion As TextBox = Me.Page.FindControl(txtCodigo.UniqueID)
        textAnotacion.Focus()
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "Select EPI_Id as Codigo, Descricao From EPI where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And EPI_Id = " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao = '" & txtDescricao.Text & "'"
        End If

        sql &= "order by EPI_Id"

        Return Banco.ConsultaDataSet(sql, "Tabelas")
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ClasseDeRisco", "GRAVAR") Then
                objClasseDeRisco = New [Lib].Negocio.ClasseDeRisco
                objClasseDeRisco.Codigo = txtCodigo.Text
                objClasseDeRisco.Descricao = txtDescricao.Text
                objClasseDeRisco.IUD = "I"
                If objClasseDeRisco.Salvar Then
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

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ClasseDeRisco", "ALTERAR") Then
                objClasseDeRisco = New [Lib].Negocio.ClasseDeRisco

                objClasseDeRisco.Codigo = gridClasseDeRisco.SelectedRow.Cells(1).Text()
                objClasseDeRisco.Descricao = txtDescricao.Text
                objClasseDeRisco.IUD = "U"
                If objClasseDeRisco.Salvar Then
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

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ClasseDeRisco", "EXCLUIR") Then
                objClasseDeRisco = New [Lib].Negocio.ClasseDeRisco

                objClasseDeRisco.Codigo = gridClasseDeRisco.SelectedRow.Cells(1).Text()
                objClasseDeRisco.IUD = "D"
                If objClasseDeRisco.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para Excluir Registro")
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
            If Funcoes.VerificaPermissao("ClasseDeRisco", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório de Classe de Risco")
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
            Funcoes.Ajuda(Me.Page, "ClasseDeRisco")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class