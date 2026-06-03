Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class IngredienteAtivo
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objIngredienteAtivo As [Lib].Negocio.IA

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("IngredienteAtivo", "ACESSAR") Then
                BuscarEstadoFisico()
                BuscarGrupoQuimico()
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissã para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BuscarEstadoFisico()
        ddl.Carregar(ddlEstadoFisico, CarregarDDL.Tabela.EstadoFisicoIA, "", True)
    End Sub

    Private Sub BuscarGrupoQuimico()
        ddl.Carregar(ddlGrupoQuimico, CarregarDDL.Tabela.GrupoQuimico, "", True)
    End Sub

    Protected Sub gridIngredienteAtivo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False

            objIngredienteAtivo = New [Lib].Negocio.IA(gridIngredienteAtivo.SelectedRow.Cells(1).Text())

            txtCodigo.Text = objIngredienteAtivo.CodigoIA
            ddlEstadoFisico.SelectedValue = objIngredienteAtivo.CodigoEstadoFisico
            ddlGrupoQuimico.SelectedValue = objIngredienteAtivo.CodigoGrupoQuimico
            txtNomeComum.Text = objIngredienteAtivo.NomeComum
            txtNomeQuimico.Text = objIngredienteAtivo.NomeQuimico
            txtSolubilidade.Text = objIngredienteAtivo.Solubilidade
            txtPesoMolecular.Text = objIngredienteAtivo.PesoMolecular
            txtPontoFusao.Text = objIngredienteAtivo.PontoFusao
            txtPressaoVapor.Text = objIngredienteAtivo.PressaoVapor

            txtNomeComum.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListIA(True)
        gridIngredienteAtivo.DataSource = Lista.ToArray()
        gridIngredienteAtivo.DataBind()
    End Sub

    Private Sub Limpar()
        objIngredienteAtivo = New [Lib].Negocio.IA
        objIngredienteAtivo.Numerador()
        txtCodigo.Text = objIngredienteAtivo.CodigoIA.ToString

        ddlEstadoFisico.SelectedIndex = 0
        ddlGrupoQuimico.SelectedIndex = 0
        txtNomeComum.Text = ""
        txtNomeQuimico.Text = ""
        txtSolubilidade.Text = ""
        txtPesoMolecular.Text = ""
        txtPontoFusao.Text = ""
        txtPressaoVapor.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtCodigo.Enabled = True

        txtCodigo.Focus()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("IngredienteAtivo", "GRAVAR") Then
                objIngredienteAtivo = New [Lib].Negocio.IA

                objIngredienteAtivo.CodigoIA = txtCodigo.Text
                objIngredienteAtivo.CodigoEstadoFisico = ddlEstadoFisico.SelectedValue
                objIngredienteAtivo.CodigoGrupoQuimico = ddlGrupoQuimico.SelectedValue
                objIngredienteAtivo.NomeComum = txtNomeComum.Text
                objIngredienteAtivo.NomeQuimico = txtNomeQuimico.Text
                objIngredienteAtivo.Solubilidade = txtSolubilidade.Text
                objIngredienteAtivo.PesoMolecular = txtPesoMolecular.Text
                objIngredienteAtivo.PontoFusao = txtPontoFusao.Text
                objIngredienteAtivo.PressaoVapor = txtPressaoVapor.Text
                objIngredienteAtivo.IUD = "I"
                If objIngredienteAtivo.Salvar Then
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
            If Funcoes.VerificaPermissao("IngredienteAtivo", "ALTERAR") Then
                objIngredienteAtivo = New [Lib].Negocio.IA

                objIngredienteAtivo.CodigoIA = txtCodigo.Text
                objIngredienteAtivo.CodigoEstadoFisico = ddlEstadoFisico.SelectedValue
                objIngredienteAtivo.CodigoGrupoQuimico = ddlGrupoQuimico.SelectedValue
                objIngredienteAtivo.NomeComum = txtNomeComum.Text
                objIngredienteAtivo.NomeQuimico = txtNomeQuimico.Text
                objIngredienteAtivo.Solubilidade = txtSolubilidade.Text
                objIngredienteAtivo.PesoMolecular = txtPesoMolecular.Text
                objIngredienteAtivo.PontoFusao = txtPontoFusao.Text
                objIngredienteAtivo.PressaoVapor = txtPressaoVapor.Text
                objIngredienteAtivo.IUD = "U"
                If objIngredienteAtivo.Salvar Then
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
            If Funcoes.VerificaPermissao("IngredienteAtivo", "ACESSAR") Then
                objIngredienteAtivo = New [Lib].Negocio.IA

                objIngredienteAtivo.CodigoIA = txtCodigo.Text
                objIngredienteAtivo.CodigoEstadoFisico = ddlEstadoFisico.SelectedValue
                objIngredienteAtivo.CodigoGrupoQuimico = ddlGrupoQuimico.SelectedValue
                objIngredienteAtivo.NomeComum = txtNomeComum.Text
                objIngredienteAtivo.NomeQuimico = txtNomeQuimico.Text
                objIngredienteAtivo.Solubilidade = txtSolubilidade.Text
                objIngredienteAtivo.PesoMolecular = txtPesoMolecular.Text
                objIngredienteAtivo.PontoFusao = txtPontoFusao.Text
                objIngredienteAtivo.PressaoVapor = txtPressaoVapor.Text
                objIngredienteAtivo.IUD = "D"
                If objIngredienteAtivo.Salvar Then
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
            If Funcoes.VerificaPermissao("IngredienteAtivo", "RELATORIO") Then
                Sql = "   Select  IA_Id as Codigo,CAST(ia.EstadoFisicoIA as varchar) + ' - ' + ef.descricao as EstadoFisico ,    " & vbCrLf & _
                "           cast(GrupoQuimico as varchar) + ' - ' + gq.Descricao as GrupoQuimico, NomeComum, NomeQuimico,  " & vbCrLf & _
                "           Solubilidade, PesoMolecular, PontoFusao, PressaoVapor                                          " & vbCrLf & _
                "   	From IA                                                                                            " & vbCrLf & _
                "   		Inner Join GrupoQuimico gq                                                                     " & vbCrLf & _
                "   			on gq.GrupoQuimico_Id = ia.GrupoQuimico                                                    " & vbCrLf & _
                "   		Inner Join EstadoFisicoIA ef                                                                   " & vbCrLf & _
                "   			on ef.EstadoFisicoIA_Id = ia.EstadoFisicoIA                                                " & vbCrLf & _
                "       Where 1=1 " & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "IngredienteAtivo")

                Funcoes.BindReport(Me.Page, ds, "Cr_IngredienteAtivo", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParameters() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(ddlEstadoFisico.SelectedValue) Then
            param &= "Estado Físico: " & ddlEstadoFisico.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlGrupoQuimico.SelectedValue) Then
            param &= "Grupo Químico: " & ddlGrupoQuimico.SelectedValue & vbCrLf
        End If

        Return param
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "IngredienteAtivo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class