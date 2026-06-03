Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class MenuDeAcesso
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("MenuDeAcesso", "ACESSAR") Then
                CarregarMenus()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Sub CarregarMenus()
        If Globais.GPermiteLeitura = "S" Then
            Sql = "  SELECT Menu_Id as Menu, Descricao " & _
                            " FROM MenuDeAcesso " & _
                            " ORDER BY Menu_Id"

            GridMenus.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridMenus.DataBind()
        End If
    End Sub

    Sub Limpar()
        txtMenu.Text = ""
        txtDescricao.Text = ""
        txtMenu.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridMenus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Limpar()
        txtMenu.Text = GridMenus.SelectedRow.Cells(1).Text()
        txtDescricao.Text = GridMenus.SelectedRow.Cells(2).Text()
        txtMenu.Enabled = False
        txtDescricao.Focus()
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Globais.GPermiteGravar = "S" Then
            Sql = "INSERT Into MenuDeAcesso (Menu_id, Descricao) "
            Sql &= " Values('" & txtMenu.Text & "' "
            Sql &= ",'" & UCase(txtDescricao.Text) & "')"

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                CarregarMenus()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        If Globais.GPermiteAlterar = "S" Then
            Sql = "UPDATE MenuDeAcesso"
            Sql &= " Set Descricao = '" & txtDescricao.Text & "' "
            Sql &= " WHERE Menu_Id = '" & txtMenu.Text & "' "
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                CarregarMenus()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If Globais.GPermiteExcluir = "S" Then
            Sql = "DELETE FROM MenuDeAcesso"
            Sql &= " WHERE Menu_Id = '" & txtMenu.Text & "' "
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                CarregarMenus()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("MenuDeAcesso", "RELATORIO") Then
                Dim Titulo As String = "Menu de Acesso"
                Dim Codigo As String = "Menu"
                Dim Descricao As String = "Descrição"

                Titulo = "Menu de Acesso"
                Sql = " Select Menu_Id as Codigo, Descricao From MenuDeAcesso Order by Menu_Id "


                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", Titulo)
                parameters.Add("Codigo", Codigo)
                parameters.Add("Descricao", Descricao)
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
            Funcoes.Ajuda(Me.Page, "MenuDeAcesso")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class