Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeAtivos
    Inherits BasePage

    Dim Sql As String

    Dim Empresa() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Patrimonio)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeAtivos", "ACESSAR") Then
                    CarregarUnidade()
                    CargaGrupos()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Patrimonio.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
            DdlEmpresa_SelectedIndexChanged(DdlEmpresa, New EventArgs)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaGrupos()
        ddl.Carregar(DdlGrupo, CarregarDDL.Tabela.GruposDeAtivos, "", True)
    End Sub

    Private Sub UltimaAtualizacao()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        Sql = " SELECT Top(1) Movimento_Id as Data" & vbCrLf & _
              "   FROM AtivosXMovimentos" & vbCrLf & _
              "  WHERE (Processo = 'NORMAL' OR Processo = 'INICIAL')" & vbCrLf & _
              "  ORDER BY Movimento_Id DESC" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "AtivosXMovimentos").Tables(0).Rows
            txtDataFinal.Text = Format(Dr("Data"), "dd/MM/yyyy")
        Next
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlEmpresa.SelectedIndexChanged
        Try
            UltimaAtualizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ckDepreciados_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ckDepreciados.Checked = True Then
                ckBaixados.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ckBaixados_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ckBaixados.Checked = True Then
                ckDepreciados.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioDeAtivos", "RELATORIO") Then
                If txtDataFinal.Text = "" Then
                    MsgBox(Me.Page, "Data da última depreciação é obrigatária.")
                Else

                    Dim Grupo As String = DdlGrupo.SelectedValue
                    Empresa = DdlEmpresa.SelectedValue.Split("-")

                    Dim Ds As New DataSet
                    Dim s As New MemoryStream
                    Dim Parametros As String = "Parametros: " & vbCrLf

                    Parametros &= "Empresa: " & DdlEmpresa.SelectedItem.Text & vbCrLf

                    Sql = " SELECT A.Empresa, A.EndEmpresa, Cli.Nome AS NomeEmpresa, Cli.Cidade, Cli.Estado, Ga.Descricao AS GrupoDescricao, " & vbCrLf & _
                          "        A.Grupo_Id AS Grupo, REPLICATE('0', 4 - LEN(CAST(A.Codigo_Id AS varchar))) + CAST(A.Codigo_Id AS varchar) AS Codigo, " & vbCrLf & _
                          "        REPLICATE('0', 2 - LEN(CAST(A.Sequencia_Id AS varchar))) + CAST(A.Sequencia_Id AS varchar) AS Sequencia, A.Historico, " & vbCrLf & _
                          "        A.CentroDeCusto, A.DataAquisicao AS Aquisicao, A.ValorOriginal, ISNULL(Am.Valor,0) AS Depreciado, A.ValorOriginal - ISNULL(Am.Valor, 0) AS Residual," & vbCrLf & _
                          " CASE " & vbCrLf & _
                          "  WHEN A.DataDaBaixa IS NOT NULL THEN A.Descricao + ' DATA DA BAIXA: ' + CONVERT(VARCHAR, A.DataDaBaixa, 103) " & vbCrLf & _
                          "  ELSE A.Descricao " & vbCrLf & _
                          "  END AS Descricao " & vbCrLf & _
                          "   FROM Ativos A " & vbCrLf & _
                          "  INNER JOIN GruposDeAtivos Ga " & vbCrLf & _
                          "     ON A.Grupo_Id = Ga.Grupo_Id " & vbCrLf & _
                          "   LEFT JOIN Clientes Cli" & vbCrLf & _
                          "     ON A.Empresa    = Cli.Cliente_Id " & vbCrLf & _
                          "    AND A.EndEmpresa = Cli.Endereco_Id" & vbCrLf & _
                          "   LEFT JOIN ( " & vbCrLf & _
                          "               SELECT SUM(Valor) AS Valor, Am.Grupo_Id, Am.Codigo_Id, Am.Sequencia_Id, Am.Empresa_Id " & vbCrLf & _
                          "                 FROM AtivosXMovimentos Am" & vbCrLf & _
                          "                WHERE Am.Movimento_Id  <=  '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                          "                GROUP BY Am.Grupo_Id, Am.Codigo_Id, Am.Sequencia_Id, Am.Empresa_Id " & vbCrLf & _
                          "             ) as Am" & vbCrLf & _
                          "     ON A.Grupo_Id     = Am.Grupo_Id" & vbCrLf & _
                          "    AND A.Codigo_Id    = Am.Codigo_Id" & vbCrLf & _
                          "    AND A.Sequencia_Id = Am.Sequencia_Id  " & vbCrLf & _
                          "    AND A.Empresa_Id   = Am.Empresa_Id" & vbCrLf & _
                          "  WHERE A." & IIf(String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue), "Empresa_Id = '" & Left(UsuarioServidor.CodigoEmpresa, 8), "Empresa = '" & Empresa(0)) & "'" & vbCrLf & _
                          "    AND A.Situacao = 1" & vbCrLf

                    If Not String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
                        Parametros &= "Grupo:" & DdlGrupo.SelectedValue & vbCrLf
                        Sql &= "   AND A.Grupo_Id ='" & Grupo & "' " & vbCrLf

                    End If
                    Parametros &= "Data da Depreciação:" & txtDataFinal.Text & vbCrLf
                    Sql &= "   AND A.DataAquisicao <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf

                    If ckBaixados.Checked Then
                        Parametros &= "Baixados " & vbCrLf
                        Sql &= "  AND (A.DataDaBaixa <= '" & txtDataFinal.Text.ToSqlDate() & "' and  A.DataDaBaixa IS not  NULL)  " & vbCrLf
                    Else
                        Sql &= "  AND (A.DataDaBaixa > '" & txtDataFinal.Text.ToSqlDate() & "' OR  A.DataDaBaixa IS NULL) " & vbCrLf
                    End If

                    If ckDepreciados.Checked Then
                        Parametros &= "Depreciados 100% " & vbCrLf
                        Sql &= "AND (A.ValorOriginal = " & vbCrLf & _
                               " ISNULL((SELECT SUM(Valor) AS Valor" & vbCrLf & _
                               "           FROM AtivosXMovimentos Am" & vbCrLf & _
                               "          WHERE A.Grupo_Id     = Am.Grupo_Id" & vbCrLf & _
                               "            AND A.Codigo_Id    = Am.Codigo_Id" & vbCrLf & _
                               "            AND A.Sequencia_Id = Am.Sequencia_Id " & vbCrLf & _
                               "            AND Am.Movimento_Id <= '" & txtDataFinal.Text.ToSqlDate() & "') , 0))" & vbCrLf
                    End If

                    If chkSeguro.Checked Then
                        Parametros &= "Seguro " & vbCrLf
                        Sql &= " AND A.Seguro = 'TRUE' " & vbCrLf
                    Else
                        Sql &= " AND A.Seguro = 'FALSE' " & vbCrLf
                    End If
                    Sql &= " AND A.InicioDeUso <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf


                    Sql &= "  --GROUP BY A.Empresa, A.EndEmpresa, A.Grupo_Id, A.Codigo_Id, A.Sequencia_Id, A.Empresa_Id, A.Descricao, A.Historico, A.CentroDeCusto, A.DataAquisicao, " & vbCrLf & _
                           "  --         A.ValorOriginal, Ga.Descricao, Ga.PercentualDepreciacao, Cli.Nome, Cli.Cidade, Cli.Estado " & vbCrLf & _
                           "  ORDER BY A.Empresa, A.EndEmpresa, Grupo, A.Codigo_Id, A.Sequencia_Id" & vbCrLf


                    Ds = Banco.ConsultaDataSet(Sql, "Ativos")
                    If Pdf Then
                        Dim Empresa As String = String.Empty

                        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                            Empresa = DdlEmpresa.SelectedValue
                        Else
                            Empresa = String.Format("{0}-{1}", UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                        End If

                        Dim obj As New Cliente(Empresa.Split("-")(0), Empresa.Split("-")(1))

                        Dim parameters = New Dictionary(Of String, Object)()
                        parameters.Add("Parametros", Parametros)

                        Funcoes.BindReport(Me.Page, Ds, "Cr_Ativos", eExportType.PDF, parameters)
                    Else
                        Dim colunas As New Dictionary(Of String, eTipoCampo)
                        colunas.Add("Aquisicao", eTipoCampo.Data)
                        colunas.Add("ValorOriginal", eTipoCampo.ValorComTotalizador)
                        colunas.Add("Depreciado", eTipoCampo.ValorComTotalizador)
                        colunas.Add("Residual", eTipoCampo.ValorComTotalizador)

                        Funcoes.BindExcelOffice(Me.Page, Ds, "Relatório de Ativos", colunas)
                    End If

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeAtivos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class