Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style

Public Class AnalisadorDeDeposito
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AnalisadorDeDeposito", "ACESSAR") Then
                    CargaUnidade()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gerencial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Try
            If ValidaCampos() Then
                Dim ds As DataSet = getDataSet()

                BindExcelOffice(Me.Page, ds, "AnalisadorDeDepositos")
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

#End Region

#Region "Methods"

    Private Function ValidaCampos() As Boolean
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a Empresa.")
                Return False
            ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                MsgBox(Me.Page, "Informe o período.")
                Return False
            End If
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Sub BindExcelOffice(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing, Optional ByVal viewAllTables As Boolean = False)
        If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 1
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                'emitir excel.xsls do office / relatório padrão em lista
                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                        For Each dt As DataTable In ds.Tables
                            Dim columnIndex As Integer = 1
                            If viewAllTables Then
                                worksheet.Cells(rowIndex, columnIndex).Value = dt.TableName
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                                worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.White)
                                rowIndex += 1
                            End If

                            'inserindo o cabeçalho
                            For Each col As DataColumn In dt.Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            Next

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            rowIndex += 1

                            'exportando conteúdo da planilha com os dados da tabela
                            For Each row As DataRow In dt.Rows
                                columnIndex = 1
                                For Each col As DataColumn In dt.Columns
                                    If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName).ToString().Replace(".", ""))
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    End If

                                    'formatações de valores
                                    If colunas IsNot Nothing Then
                                        For Each coluna In colunas
                                            If coluna.Key = col.ColumnName Then
                                                If coluna.Value = eTipoCampo.Numerico Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.ValorComTotalizador OrElse coluna.Value = eTipoCampo.ValorSemTotalizador Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.Data Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    Else
                                        If IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                        ElseIf IsDate(row(col.ColumnName)) Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                        End If
                                    End If
                                    columnIndex += 1
                                Next

                                'formatações de celulas
                                If row("OP") = 55 And DdlEmpresa.SelectedValue.Split("-")(0) = row("Deposito") Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                        range.Style.Font.Bold = True
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Yellow)
                                        range.Style.Font.Color.SetColor(Color.Red)
                                    End Using
                                ElseIf rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If

                                rowIndex += 1
                            Next

                            'aplicando formatação nas células do rodapé
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            columnIndex = 1

                            'soma dos campos de valores
                            If colunas IsNot Nothing Then

                                For Each col In colunas
                                    If col.Value = eTipoCampo.ValorComTotalizador Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                                        Exit For
                                    End If
                                Next

                                For Each col As DataColumn In dt.Columns
                                    For Each coluna In colunas
                                        If coluna.Key = col.ColumnName Then
                                            If coluna.Value = eTipoCampo.ValorComTotalizador Then
                                                worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(1, columnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, columnIndex).Address & ")"
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    columnIndex += 1
                                Next
                            End If

                            If Not viewAllTables Then
                                Exit For
                            End If

                            rowIndex += 2
                        Next

                        'criando auto filtro na planilha
                        If Not viewAllTables Then
                            worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count).AutoFilter = True
                        End If

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linham
                        worksheet.View.FreezePanes(2, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

    Private Sub CargaUnidade()
        Try
            ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Try
            Dim sql As String = String.Empty

            sql = " SELECT NotasFiscaisXItens.Produto_Id As Produto, Produtos.Nome, NotasFiscais.Deposito, Deposito.Nome AS NomeDoDeposito, NotasFiscais.Operacao as OP, NotasFiscais.SubOperacao as SO, SubOperacoes.EstoqueFisico AS Fisico," & vbCrLf &
                  "           SubOperacoes.Descricao, NotasFiscaisXItens.EntradaSaida_Id As ES,  " & vbCrLf &
                  "       SUM(Case when NotasFiscaisXItens.EntradaSaida_ID = 'E' then NotasFiscaisXItens.PesoFiscal else 0 End) AS Entradas, " & vbCrLf &
                  "       SUM(Case when NotasFiscaisXItens.EntradaSaida_ID = 'S' then NotasFiscaisXItens.PesoFiscal else 0 End) AS Saidas " & vbCrLf &
                  " FROM  NotasFiscais INNER JOIN" & vbCrLf &
                  "           NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                  "           NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                  "           NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                  "           NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                  "           Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id INNER JOIN" & vbCrLf &
                  "           SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
                  "           Clientes AS Empresa ON NotasFiscais.Empresa_Id = Empresa.Cliente_Id AND NotasFiscais.EndEmpresa_Id = Empresa.Endereco_Id INNER JOIN" & vbCrLf &
                  "           Clientes AS Deposito ON NotasFiscais.Deposito = Deposito.Cliente_Id AND NotasFiscais.EndDeposito = Deposito.Endereco_Id" & vbCrLf &
                  " WHERE (NotasFiscais.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "') And (NotasFiscais.Situacao = 1)" & vbCrLf &
                  "   And (NotasFiscais.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
                  "   And (Produtos.Agrupar = 'N') AND (NotasFiscais.Operacao < 68)" & vbCrLf &
                  " GROUP BY NotasFiscaisXItens.Produto_Id, NotasFiscais.Operacao, NotasFiscais.SubOperacao, NotasFiscais.Deposito, NotasFiscais.Empresa_Id," & vbCrLf &
                  "           SubOperacoes.Descricao , NotasFiscaisXItens.EntradaSaida_Id, Produtos.Nome, Empresa.Nome, Deposito.Nome, SubOperacoes.EstoqueFisico" & vbCrLf &
                  " ORDER BY NotasFiscaisXItens.Produto_Id, NotasFiscais.Deposito, NotasFiscais.Operacao, NotasFiscais.SubOperacao" & vbCrLf

            Return Banco.ConsultaDataSet(sql, "AnalisadorDeDepositos")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    Private Sub Limpar()
        txtDataInicial.Text = String.Format("01/01/{0}", Year(Now))
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        LiberaEmpresa()
    End Sub
    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AnalisadorDeDeposito")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class