Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Bancos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Bancos", "ACESSAR") Then
                CarregarBancos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarBancos(Optional ByVal codigo As Integer = 0)
        Dim ds As New DataSet()
        Sql = "  SELECT Banco_id as Codigo, Descricao, Ativo, isnull(LiquidacaoDias,0) as LiquidacaoDias" & vbCrLf & _
                " FROM Bancos " & vbCrLf
        If codigo > 0 Then
            Sql &= "Where Banco_Id = " & codigo
        End If
        Sql &= " ORDER BY Banco_id"

        ds = Banco.ConsultaDataSet(Sql, "Consulta")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            GridBancos.DataSource = ds
            GridBancos.DataBind()
        Else
            MsgBox(Me.Page, "Nenhum resultado encontrado.")
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = String.Empty
        txtDescricao.Text = String.Empty
        txtCodigo.Enabled = True
        txtLiquidacaoDias.Text = String.Empty
        CarregarBancos()

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function Existe(ByVal codigo As Integer) As Boolean
        Dim sql As String = " Select * from Bancos Where Banco_Id = " & codigo
        Dim ds As New DataSet()
        ds = Banco.ConsultaDataSet(sql, "Consulta")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        End If

        Return False
    End Function

    Protected Sub GridBancos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridBancos.SelectedIndexChanged
        Try
            Limpar()
            txtCodigo.Text = GridBancos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridBancos.SelectedRow.Cells(2).Text()
            txtLiquidacaoDias.Text = GridBancos.SelectedRow.Cells(3).Text()

            If GridBancos.SelectedRow.Cells(3).Text() = "S" Then
                RadSim.Checked = True
                RadNao.Checked = False
            Else
                RadNao.Checked = True
                RadSim.Checked = False
            End If

            txtCodigo.Enabled = False
            txtDescricao.Focus()

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                CarregarBancos(txtCodigo.Text)
            Else
                CarregarBancos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Bancos", "GRAVAR") Then
                If Not String.IsNullOrWhiteSpace(txtCodigo.Text) AndAlso Not Existe(txtCodigo.Text) Then
                    Sql = "INSERT Into Bancos (Banco_id, Descricao, Ativo, LiquidacaoDias) " & vbCrLf & _
                          " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                          ",'" & UCase(txtDescricao.Text) & "' " & vbCrLf & _
                          ",'" & IIf(RadSim.Checked, "S", "N") & "' " & vbCrLf & _
                          ", " & IIf(String.IsNullOrWhiteSpace(txtLiquidacaoDias.Text), 0, txtLiquidacaoDias.Text) & ")" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarBancos()
                    End If
                Else
                    MsgBox(Me.Page, "Código já está cadastrado.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Bancos", "ALTERAR") Then
                Sql = "UPDATE Bancos" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      " , Ativo = '" & IIf(RadSim.Checked, "S", "N") & "'" & vbCrLf & _
                      " , LiquidacaoDias = '" & IIf(String.IsNullOrWhiteSpace(txtLiquidacaoDias.Text) OrElse Not IsNumeric(txtLiquidacaoDias.Text), 0, txtLiquidacaoDias.Text) & "'" & vbCrLf & _
                      " WHERE Banco_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarBancos()
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
            If Funcoes.VerificaPermissao("Bancos", "EXCLUIR") Then
                Sql = "DELETE FROM Bancos" & vbCrLf & _
                      " WHERE Banco_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarBancos()
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
            If Funcoes.VerificaPermissao("Bancos", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 4 - LEN(CAST(Banco_Id AS varchar))) + CAST(Banco_Id AS varchar) as Codigo, Descricao From Bancos Order by Banco_Id"
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Bancos.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

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
            Funcoes.Ajuda(Me.Page, "Bancos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class