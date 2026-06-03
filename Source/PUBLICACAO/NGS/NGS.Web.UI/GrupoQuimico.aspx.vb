Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class GrupoQuimico
    Inherits BasePage

    Dim objGrupoQuimico As [Lib].Negocio.GrupoQuimico

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("GrupoQuimico", "ACESSAR") Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridGrupoQuimico_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = True
            txtCodigo.Text = Server.HtmlDecode(gridGrupoQuimico.SelectedRow.Cells(1).Text())
            txtDescricao.Text = Server.HtmlDecode(gridGrupoQuimico.SelectedRow.Cells(2).Text())
            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListGrupoQuimico(True)
        gridGrupoQuimico.DataSource = Lista.ToArray()
        gridGrupoQuimico.DataBind()
    End Sub

    Private Sub Limpar()
        objGrupoQuimico = New [Lib].Negocio.GrupoQuimico
        objGrupoQuimico.VerSequencia()
        txtCodigo.Text = objGrupoQuimico.Codigo
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtCodigo.Focus()
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "Select GrupoQuimico_Id AS Codigo, Descricao From GrupoQuimico where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And GrupoQuimico_Id = " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao = '" & txtDescricao.Text & "'"
        End If

        sql &= "order by GrupoQuimico_Id"

        Return Banco.ConsultaDataSet(sql, "Tabelas")
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("GrupoQuimico", "GRAVAR") Then
                objGrupoQuimico = New [Lib].Negocio.GrupoQuimico
                objGrupoQuimico.Codigo = txtCodigo.Text
                objGrupoQuimico.Descricao = txtDescricao.Text
                objGrupoQuimico.IUD = "I"
                If objGrupoQuimico.Salvar Then
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
            If Funcoes.VerificaPermissao("GrupoQuimico", "ALTERAR") Then
                objGrupoQuimico = New [Lib].Negocio.GrupoQuimico
                objGrupoQuimico.Codigo = gridGrupoQuimico.SelectedRow.Cells(1).Text()
                objGrupoQuimico.Descricao = txtDescricao.Text
                objGrupoQuimico.IUD = "U"
                If objGrupoQuimico.Salvar Then
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
            If Funcoes.VerificaPermissao("GrupoQuimico", "EXCLUIR") Then
                objGrupoQuimico = New [Lib].Negocio.GrupoQuimico
                objGrupoQuimico.Codigo = gridGrupoQuimico.SelectedRow.Cells(1).Text()
                objGrupoQuimico.IUD = "D"
                If objGrupoQuimico.Salvar Then
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
            If Funcoes.VerificaPermissao("GrupoQuimico", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Grupo Químico")
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
            Funcoes.Ajuda(Me.Page, "GrupoQuimico")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class