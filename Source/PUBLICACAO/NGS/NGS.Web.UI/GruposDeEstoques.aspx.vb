Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class GruposDeEstoques
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("GruposDeEstoques", "ACESSAR") Then
                CarregarGruposDeEstoques()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarGruposDeEstoques()
        Try
            If Funcoes.VerificaPermissao("GruposDeEstoques", "LEITURA") Then
                Dim Grupo As New [Lib].Negocio.ListGrupoProduto(True, "", "", "")
                Session("ssGrupo") = Grupo
                GridGruposDeEstoques.DataSource = Grupo.ToArray
                GridGruposDeEstoques.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        chkCusto.Checked = False
        chkAgrupaFinanceiro.Checked = False
        chkRelatorioEstoque.Checked = False
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridGruposDeEstoques_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtCodigo.Text = CType(Session("ssGrupo"), [Lib].Negocio.ListGrupoProduto).Item(GridGruposDeEstoques.SelectedIndex).Codigo
            txtDescricao.Text = CType(Session("ssGrupo"), [Lib].Negocio.ListGrupoProduto).Item(GridGruposDeEstoques.SelectedIndex).Descricao
            chkCusto.Checked = CType(Session("ssGrupo"), [Lib].Negocio.ListGrupoProduto).Item(GridGruposDeEstoques.SelectedIndex).Custo
            chkAgrupaFinanceiro.Checked = CType(Session("ssGrupo"), [Lib].Negocio.ListGrupoProduto).Item(GridGruposDeEstoques.SelectedIndex).AgrupaFinanceiro
            chkRelatorioEstoque.Checked = CType(Session("ssGrupo"), [Lib].Negocio.ListGrupoProduto).Item(GridGruposDeEstoques.SelectedIndex).RelatorioEstoque
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
        If Funcoes.VerificaPermissao("GruposDeEstoques", "GRAVAR") Then
            If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                MsgBox(Me.Page, "Código não foi informado.")
                Exit Sub
            ElseIf Len(txtCodigo.Text) < 5 Then
                chkRelatorioEstoque.Checked = False
            End If

            Dim descricao As String = Funcoes.EliminarCaracteresEspeciais(txtDescricao.Text)

            Sql = "INSERT Into GruposDeEstoques(Grupo_id, Descricao, Custo, AgruparFinanceiro, MapaDeEstoque) " & vbCrLf & _
              " Values('" & txtCodigo.Text & "' " & vbCrLf & _
              ",'" & descricao.ToUpper & "', " & vbCrLf & _
              IIf(chkCusto.Checked, "1", "0") & ", " & vbCrLf & _
              IIf(chkAgrupaFinanceiro.Checked, "1", "0") & ", " & vbCrLf & _
              IIf(chkRelatorioEstoque.Checked, "1", "0") & ")" & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) Then
                Limpar()
                CarregarGruposDeEstoques()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("GruposDeEstoques", "ALTERAR") Then
                If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                    MsgBox(Me.Page, "Código não foi informado.")
                    Exit Sub
                ElseIf Len(txtCodigo.Text) < 5 Then
                    chkRelatorioEstoque.Checked = False
                End If

                Dim descricao As String = Funcoes.EliminarCaracteresEspeciais(txtDescricao.Text)

                Sql = "UPDATE GruposDeEstoques" & vbCrLf & _
                      " Set Descricao = '" & descricao.ToUpper & "', " & vbCrLf & _
                      "Custo = " & IIf(chkCusto.Checked, "1", "0") & ", " & vbCrLf & _
                      "AgruparFinanceiro = " & IIf(chkAgrupaFinanceiro.Checked, "1", "0") & ", " & vbCrLf & _
                      "MapaDeEstoque = " & IIf(chkRelatorioEstoque.Checked, "1", "0") & vbCrLf & _
                      " WHERE Grupo_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CarregarGruposDeEstoques()
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
            If Funcoes.VerificaPermissao("GruposDeEstoques", "EXCLUIR") Then
                Sql = "DELETE FROM GruposDeEstoques" & vbCrLf & _
                      " WHERE Grupo_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CarregarGruposDeEstoques()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para excluir registro.")
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
            If Funcoes.VerificaPermissao("GruposDeEstoques", "RELATORIO") Then
                Sql = "Select Grupo_Id AS Codigo, Descricao, MapaDeEstoque From GruposDeEstoques Order by Grupo_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Grupos De Estoques")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
                'Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                'parameters.Add("Empresa", objEmpresa.Nome.Trim())
                'parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

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
            Funcoes.Ajuda(Me.Page, "GruposDeEstoques")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class