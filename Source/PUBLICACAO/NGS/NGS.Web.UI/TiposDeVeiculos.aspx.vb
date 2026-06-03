Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TiposDeVeiculos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Private objTipoDeVeiculo As [Lib].Negocio.TipoDeVeiculo

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TiposDeVeiculos", "ACESSAR") Then
                CarregarTiposDeVeiculos()
                ddl.Carregar(ddlViaTransporte, CarregarDDL.Tabela.ViaDeTransportes, "")
                txtDescricao.Focus()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarTiposDeVeiculos()
        If Funcoes.VerificaPermissao("TiposDeVeiculos", "LEITURA") Then
            Dim ListTiposDeVeiculos = New TipoDeVeiculos(True)
            GridTiposDeVeiculos.DataSource = ListTiposDeVeiculos.ToArray
            GridTiposDeVeiculos.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCapacidade.Text = ""
        txtTaraMinima.Text = ""
        txtTaraMaxima.Text = ""
        ddlViaTransporte.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridTiposDeVeiculos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            With GridTiposDeVeiculos.SelectedRow
                txtCodigo.Text = .Cells(1).Text
                txtDescricao.Text = .Cells(2).Text
                txtCapacidade.Text = .Cells(3).Text.Replace(".", "")
                txtTaraMinima.Text = .Cells(4).Text.Replace(".", "")
                txtTaraMaxima.Text = .Cells(5).Text.Replace(".", "")

                Dim txtCodigoVia As HiddenField = CType(.Cells(6).FindControl("txtCodigoVia"), HiddenField)
                Dim itemSelecionado As ListItem = ddlViaTransporte.Items.FindByValue(txtCodigoVia.Value)
                If itemSelecionado Is Nothing Then
                    ddlViaTransporte.SelectedIndex = 0
                Else
                    ddlViaTransporte.SelectedIndex = -1
                    itemSelecionado.Selected = True
                End If
            End With
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeVeiculos", "GRAVAR") Then
                If ValidarCampos() Then
                    objTipoDeVeiculo = New TipoDeVeiculo()
                    objTipoDeVeiculo.IUD = "I"

                    objTipoDeVeiculo.Descricao = txtDescricao.Text
                    objTipoDeVeiculo.Capacidade = txtCapacidade.Text
                    If ddlViaTransporte.SelectedValue = "0" Or String.IsNullOrWhiteSpace(ddlViaTransporte.SelectedValue) Then
                        objTipoDeVeiculo.ViaDeTransporte = Nothing
                    Else
                        objTipoDeVeiculo.ViaDeTransporte = ddlViaTransporte.SelectedValue
                    End If
                    objTipoDeVeiculo.TaraMinima = txtTaraMinima.Text
                    objTipoDeVeiculo.TaraMaxima = txtTaraMaxima.Text

                    If objTipoDeVeiculo.Salvar Then
                        Limpar()
                        CarregarTiposDeVeiculos()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
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
            If Funcoes.VerificaPermissao("TiposDeVeiculos", "ALTERAR") Then
                If ValidarCampos() Then
                    objTipoDeVeiculo = New TipoDeVeiculo(txtCodigo.Text)
                    objTipoDeVeiculo.IUD = "U"
                    'objTipoDeVeiculo.Codigo = txtCodigo.Text
                    objTipoDeVeiculo.Descricao = txtDescricao.Text
                    objTipoDeVeiculo.Capacidade = txtCapacidade.Text
                    If ddlViaTransporte.SelectedValue = "0" Or String.IsNullOrWhiteSpace(ddlViaTransporte.SelectedValue) Then
                        objTipoDeVeiculo.ViaDeTransporte = Nothing
                    Else
                        objTipoDeVeiculo.ViaDeTransporte = ddlViaTransporte.SelectedValue
                    End If
                    objTipoDeVeiculo.TaraMinima = txtTaraMinima.Text
                    objTipoDeVeiculo.TaraMaxima = txtTaraMaxima.Text

                    If objTipoDeVeiculo.Salvar Then
                        Limpar()
                        CarregarTiposDeVeiculos()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
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
            If Funcoes.VerificaPermissao("TiposDeVeiculos", "EXCLUIR") Then
                objTipoDeVeiculo = New TipoDeVeiculo(txtCodigo.Text)
                objTipoDeVeiculo.IUD = "D"
                If objTipoDeVeiculo.Salvar Then
                    Limpar()
                    CarregarTiposDeVeiculos()
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeVeiculos", "RELATORIO") Then
                Sql = "SELECT REPLICATE('0', 3 - LEN(CAST(TV.Codigo_Id AS varchar))) + CAST(TV.Codigo_Id AS varchar) as Codigo_Id, " & vbCrLf & _
                      "TV.Descricao, Capacidade, TaraMinima, TaraMaxima, VT.Descricao AS ViaDeTransporte " & vbCrLf & _
                      "FROM  TiposDeVeiculos TV " & vbCrLf & _
                      "LEFT JOIN ViaDeTransportes VT " & vbCrLf & _
                      "ON VT.Codigo_Id = TV.ViaDeTransporte " & vbCrLf

                DS = Banco.ConsultaDataSet(Sql, "TiposDeVeiculos")

                Funcoes.BindReport(Me.Page, DS, "Cr_TiposDeVeiculos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TiposDeVeiculos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição do tipo de veículo não foi informada.")
            txtDescricao.Focus()
            Return False
        ElseIf ddlViaTransporte.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Via de transporte não foi informada.")
            ddlViaTransporte.Focus()
            Return False
        Else
            Return True
        End If
    End Function

End Class