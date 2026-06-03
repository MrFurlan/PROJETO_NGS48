Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Numerador
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Cliente As String
    Dim campo() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Numerador", "ACESSAR") Then
                    CargaUnidade()
                    VerificaUnidade()
                    Limpar()
                    CarregarNumeradores()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarNumeradores()
        If Funcoes.VerificaPermissao("Numerador", "LEITURA") Then
            Sql = "  SELECT  Numerador_ID as Numerador, Descricao, Sequencia, isnull(Serie, '') as Serie" & vbCrLf & _
                  " FROM Numerador " & vbCrLf & _
                  " Where Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' And EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                  " ORDER BY Numerador_id" & vbCrLf
            GridNumeradores.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridNumeradores.DataBind()
        End If
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub Limpar()
        txtNumerador.Text = ""
        txtDescricao.Text = ""
        txtSequencia.Text = ""
        txtSerie.Text = ""

        txtNumerador.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        LiberaEmpresa()
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(DdlUnidade.SelectedValue) Then
            param &= "Unidade: " & DdlUnidade.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            param &= "Empresa: " & DdlEmpresa.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtNumerador.Text) Then
            param &= "Numerador: " & txtNumerador.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "" & txtDescricao.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtSerie.Text) Then
            param &= "Série: " & txtSerie.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtSequencia.Text) Then
            param &= "Sequência: " & txtSequencia.Text
        End If

        Return param
    End Function

    Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                Exit Sub
            End If
            CarregarNumeradores()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridNumeradores_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtNumerador.Text = GridNumeradores.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridNumeradores.SelectedRow.Cells(2).Text()
            txtSerie.Text = Server.HtmlDecode(GridNumeradores.SelectedRow.Cells(3).Text())
            txtSequencia.Text = Server.HtmlDecode(GridNumeradores.SelectedRow.Cells(4).Text())
            txtNumerador.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = True
            lnkAtualizar.Parent.Visible = True


        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe à empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNumerador.Text) Then
            MsgBox(Me.Page, "Informe o numerador.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe à descricao.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtSequencia.Text) Then
            MsgBox(Me.Page, "Informe à sequencia.")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Numerador", "GRAVAR") Then
                If Validar() Then
                    Sql = "INSERT Into Numerador (Empresa_Id, EndEmpresa_Id, Numerador_id, Descricao, Sequencia, Serie) " & vbCrLf & _
                                          " Values('" & DdlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf & _
                                          "," & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                                          ", " & CInt(txtNumerador.Text) & vbCrLf & _
                                          ", '" & UCase(txtDescricao.Text) & "' " & vbCrLf & _
                                          ", " & CInt(txtSequencia.Text) & vbCrLf & _
                                          ", '" & UCase(txtSerie.Text) & "')" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarNumeradores()
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
            If Funcoes.VerificaPermissao("Numerador", "ALTERAR") Then
                If Validar() Then
                    Sql = " UPDATE Numerador" & vbCrLf & _
                                          "    Set Descricao = '" & UCase(txtDescricao.Text) & "' " & vbCrLf & _
                                          "       ,Sequencia = " & CInt(txtSequencia.Text) & vbCrLf & _
                                          "       ,Serie = '" & txtSerie.Text & "'" & vbCrLf & _
                                          " WHERE Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                          " and EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                                          " and Numerador_Id = " & CInt(txtNumerador.Text) & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarNumeradores()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão parar alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Numerador", "EXCLUIR") Then
                Sql = "DELETE FROM Numerador" & vbCrLf & _
                      " WHERE Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                      "   AND Numerador_Id = " & CInt(txtNumerador.Text) & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarNumeradores()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para exluir registro.")
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
            If Funcoes.VerificaPermissao("Numerador", "RELATORIO") Then
                Sql = " SELECT     Numerador.Empresa_Id, Numerador.EndEmpresa_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Clientes.Reduzido, Numerador.Numerador_Id, " & vbCrLf & _
                      " Numerador.Descricao, Numerador.Sequencia" & vbCrLf & _
                      " FROM         Numerador INNER JOIN" & vbCrLf & _
                      " Clientes ON Numerador.Empresa_Id = Clientes.Cliente_Id AND Numerador.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Numerador")

                Funcoes.BindReport(Me.Page, DS, "Cr_Numerador", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Numerador")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class