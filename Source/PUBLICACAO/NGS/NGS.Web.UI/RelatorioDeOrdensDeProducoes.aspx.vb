Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports CrystalDecisions.CrystalReports.Engine

Public Class RelatorioDeOrdensDeProducoes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioDeOrdensDeProducoes", "ACESSAR") Then

                CarregarUnidade()
                Limpar()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub lnkRelatorioPDF_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorioPDF.Click

        If ValidarReport() Then

            Dim crptRelatorio As New ReportDocument()

            Try
                Dim objBanco As New AcessaBanco()

                Dim strSQL As String = sqlReportPDF()

                Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "OrdemDeProducao")

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Sem registros para esse período.", eTitulo.Info)
                    Exit Sub
                End If

                Dim objEmpresa As New Cliente()

                If DdlEmpresa.SelectedIndex > 0 Then
                    objEmpresa = New Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))
                Else
                    objEmpresa = New Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
                End If

                Dim dtImagem As DataTable = ds.Tables.Add("Images")
                dtImagem.Columns.Add("path", GetType(String))
                dtImagem.Columns.Add("image", GetType(System.Byte()))
                dtImagem.Columns.Add("empresa", GetType(String))
                dtImagem.Columns.Add("cidadeEstado", GetType(String))
                dtImagem.Columns.Add("usuario", GetType(String))


                Dim drImagem As DataRow = dtImagem.NewRow()
                Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & objEmpresa.Imagem)

                drImagem("path") = strCaminhoImagem
                drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                drImagem("empresa") = objEmpresa.Nome
                drImagem("cidadeEstado") = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.Estado.Codigo)
                drImagem("usuario") = Session("ssNomeUsuario")
                dtImagem.Rows.Add(drImagem)

                Dim strCaminho As String = "~/Reports/Cr_OrdensDeProducoes.rpt"

                strCaminho = HttpContext.Current.Server.MapPath(strCaminho)
                crptRelatorio.Load(strCaminho)

                Dim strNomeArquivo As String = String.Empty
                'If pdf Then
                strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                'Else
                'strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
                'End If

                Dim strArquivo As String = HttpContext.Current.Server.MapPath(strNomeArquivo)

                crptRelatorio.SetDataSource(ds)

                If Dir(strArquivo).Length > 0 Then Kill(strArquivo)

                'If pdf Then
                crptRelatorio.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, strArquivo)
                'Else
                'crptRelatorio.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, strArquivo)
                'End If

                If IO.File.Exists(strArquivo) Then
                    'If pdf Then
                    Funcoes.AbrirArquivo(Page, strNomeArquivo)
                    'Else
                    'Funcoes.AbrirExcel(Page, strNomeArquivo)
                    'End If
                End If

            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crptRelatorio.Close()
                crptRelatorio.Dispose()
            End Try

        End If

    End Sub

    Protected Sub lnkRelatorioExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorioExcel.Click

        If ValidarReport() Then

            Dim crptRelatorio As New ReportDocument()

            Try
                Dim objBanco As New AcessaBanco()

                Dim strSQL As String = sqlReportExcel()

                Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "OrdemDeProducao")

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Sem registros para esse período.", eTitulo.Info)
                    Exit Sub
                End If

                Dim objEmpresa As New Cliente()

                If DdlEmpresa.SelectedIndex > 0 Then
                    objEmpresa = New Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))
                Else
                    objEmpresa = New Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
                End If

                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando planilha títulos
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Resumo")

                        'criando linha com o cabeçalho da planilha
                        Dim rowIndex As Integer = 8
                        Dim columnIndex As Integer = 1

                        worksheet.Cells("A1:F1").Merge = True
                        worksheet.Cells("A2:F2").Merge = True
                        worksheet.Cells("A3:F3").Merge = True
                        worksheet.Cells("A4:F4").Merge = True
                        worksheet.Cells("A5:F5").Merge = True
                        worksheet.Cells("A6:F6").Merge = True
                        worksheet.Cells("A7:F7").Merge = True
                        worksheet.Cells("A1:F7").Style.Font.Bold = True
                        worksheet.Cells("A1:F7").Style.Font.Size = 14

                        worksheet.Cells("A1").Value = objEmpresa.Nome
                        worksheet.Cells("A2").Value = "           " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                        worksheet.Cells("A3").Value = "           CNPJ - " & objEmpresa.CodigoFormatado

                        'criando linha que informa o título do relatório
                        worksheet.Cells("A5").Style.Font.Bold = True
                        worksheet.Cells("A5").Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells("A5").Value = "           RELATÓRIO DAS ORDENS DE PRODUÇÃO"

                        worksheet.Cells("A7").Value = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables(0).Columns
                            If col.ColumnName = "UsuarioInclusao" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "USUÁRIO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "UsuarioInclusaoData" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "DATA INCLUSÃO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "UsuarioAlteracao" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "USUÁRIO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "UsuarioAlteracaoData" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "DATA ALTERAÇÃO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "UsuarioCancelamento" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "USUÁRIO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "UsuarioCancelamentoData" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "DATA CANCELAMENTO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "QuantidadeDeAjuste" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "AJUSTE"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ProdutoNome" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "NOME"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ProdutoConsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "PRODUTO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ProdutoNomeConsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "NOME"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "LoteConsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "LOTE"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "QuantidadeConsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "QUANTIDADE"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ValidadeConsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "VALIDADE"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ProdutoInsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "PRODUTO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ProdutoNomeInsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "NOME"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "QuantidadeInsumo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "QUANTIDADE"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Sequencia" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "ITEM"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName.ToUpper
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            End If

                            columnIndex += 1
                        Next

                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        rowIndex += 1

                        ' criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                columnIndex += 1
                            Next

                            'aplicando formatação nas células do conteúdo
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando quinta linha (cabeçalho)
                        worksheet.View.FreezePanes(8, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crptRelatorio.Close()
                crptRelatorio.Dispose()
            End Try

        End If

    End Sub

    Private Function ValidarReport() As Boolean

        If Funcoes.VerificaPermissao("OrdemDeProducao", "RELATORIO") Then
            If ddlUnidade.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Unidade de Negócio não foi selecionada.", eTitulo.Info)
                Return False
            ElseIf DdlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
                Return False
            ElseIf txtDataInicial.Text.Length = 0 Then
                MsgBox(Me.Page, "Informe a data inicial.", eTitulo.Info)
                Return False
            ElseIf txtDataFinal.Text.Length = 0 Then
                MsgBox(Me.Page, "Informe a data final.", eTitulo.Info)
                Return False
            ElseIf CType(txtDataFinal.Text, Date) < CType(txtDataInicial.Text, Date) Then
                MsgBox(Me.Page, "Data final não pode ser menor que data inicial.", eTitulo.Info)
                Return False
            Else
                Return True
            End If

        Else

            MsgBox(Me.Page, "Usuário sem permissão para tirar Relatório", eTitulo.Info)
            Return False

        End If

    End Function

    Private Function sqlReportExcel() As String

        Dim strSQL As String = "SELECT o.Empresa_Id, o.EndEmpresa_Id, EMP.Nome + ' - ' + EMP.Reduzido AS Empresa, o.Ordem_Id AS Ordem, OPXP.Produto_Id AS Produto, p.Nome AS ProdutoNome, OPXP.Lote, convert(varchar,o.Movimento,103) AS Movimento, convert(varchar,o.Validade,103) AS Validade, OPXP.Quantidade, OPXP.QuantidadeDeAjuste, " & vbCrLf &
                                                    "		case" & vbCrLf &
                                                    "			when o.Estoque = 0" & vbCrLf &
                                                    "				then 'NÃO'" & vbCrLf &
                                                    "				else 'SIM'" & vbCrLf &
                                                    "			end AS Estoque," & vbCrLf &
                                                    "		o.UsuarioInclusao, convert(varchar,o.UsuarioInclusaoData,103) AS UsuarioInclusaoData, " & vbCrLf &
                                                    "		isnull(o.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf &
                                                    "		case" & vbCrLf &
                                                    "			when isnull(o.UsuarioAlteracaoData,'') = ''" & vbCrLf &
                                                    "				then ''" & vbCrLf &
                                                    "				else convert(varchar,o.UsuarioAlteracaoData,103)" & vbCrLf &
                                                    "			end AS UsuarioAlteracaoData," & vbCrLf &
                                                    "		isnull(o.UsuarioCancelamento,'') AS UsuarioCancelamento, " & vbCrLf &
                                                    "		case" & vbCrLf &
                                                    "			when isnull(o.UsuarioCancelamentoData,'') = ''" & vbCrLf &
                                                    "				then ''" & vbCrLf &
                                                    "				else convert(varchar,o.UsuarioCancelamentoData,103) " & vbCrLf &
                                                    "			end AS UsuarioCancelamentoData, " & vbCrLf &
                                                    "			ROW_NUMBER() OVER (PARTITION BY  EMP.Reduzido, o.Ordem_Id ORDER BY EMP.Reduzido, o.Ordem_Id, opXL.Produto_Id) AS Sequencia, " & vbCrLf &
                                                    "			opXL.Produto_Id As ProdutoConsumo, " & vbCrLf &
                                                    "			pCons.Nome As ProdutoNomeConsumo, " & vbCrLf &
                                                    "			opXL.Lote_Id As LoteConsumo, " & vbCrLf &
                                                    "			opXL.Quantidade AS QuantidadeConsumo, " & vbCrLf &
                                                    "			CONVERT(VARCHAR,opXL.Validade,103) AS ValidadeConsumo, " & vbCrLf &
                                                    "			ISNULL(pIns.Produto_Id, '') As ProdutoInsumo, " & vbCrLf &
                                                    "			ISNULL(pIns.Nome, '') As ProdutoNomeInsumo, " & vbCrLf &
                                                    "			ISNULL(opXI.Quantidade, 0) AS QuantidadeInsumo " & vbCrLf &
                                                    "FROM OrdemDeProducao o " & vbCrLf &
                                                    "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                                                    "   ON o.Empresa_Id		            = OPXP.Empresa_Id " & vbCrLf &
                                                    "   AND o.EndEmpresa_Id	            = OPXP.EndEmpresa_Id " & vbCrLf &
                                                    "   AND o.Ordem_Id			        = OPXP.Ordem_Id " & vbCrLf &
                                                    "INNER JOIN OrdemDeProducaoXConsumo opXC " & vbCrLf &
                                                    "		ON o.Empresa_id             = opXC.Empresa_id" & vbCrLf &
                                                    "		AND o.EndEmpresa_Id         = opXC.EndEmpresa_Id" & vbCrLf &
                                                    "		AND o.Ordem_Id              = opXC.Ordem_Id" & vbCrLf &
                                                    "		AND OPXP.Produto_Id         = opXC.ProdutoProducao_Id" & vbCrLf &
                                                    "INNER JOIN Clientes As EMP " & vbCrLf &
                                                    "		ON o.Empresa_Id				= EMP.Cliente_Id " & vbCrLf &
                                                    "		AND o.EndEmpresa_Id			= EMP.Endereco_Id" & vbCrLf &
                                                    "INNER JOIN Produtos p" & vbCrLf &
                                                    "		ON p.Produto_id             = OPXP.Produto_Id" & vbCrLf &
                                                    "INNER JOIN OrdemDeProducaoXConsumoXLote opXL" & vbCrLf &
                                                    "		ON o.Empresa_id             = opXL.Empresa_id" & vbCrLf &
                                                    "		AND o.EndEmpresa_Id         = opXL.EndEmpresa_Id" & vbCrLf &
                                                    "		AND o.Ordem_Id              = opXL.Ordem_Id" & vbCrLf &
                                                    "		AND opXC.Produto_Id         = opXL.Produto_Id" & vbCrLf &
                                                    "INNER JOIN Produtos pCons" & vbCrLf &
                                                    "		ON opXL.Produto_Id = pCons.Produto_Id" & vbCrLf &
                                                    "LEFT JOIN OrdemDeProducaoXInsumo opXI " & vbCrLf &
                                                    "	    ON o.Empresa_id             = opXI.Empresa_id " & vbCrLf &
                                                    "	    AND o.EndEmpresa_Id         = opXI.EndEmpresa_Id " & vbCrLf &
                                                    "	    AND o.Ordem_Id              = opXI.Ordem_Id " & vbCrLf &
                                                    "		AND OPXP.Produto_Id         = opXI.Produto_Id" & vbCrLf &
                                                    "LEFT JOIN Produtos pIns " & vbCrLf &
                                                    "		ON opXI.Produto_Id = pIns.Produto_Id " & vbCrLf &
                                                    "WHERE o.Empresa_Id    = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                                    "       AND o.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

        If rbCancelada.Checked Then
            strSQL &= "AND o.Situacao = 3" & vbCrLf
        Else
            strSQL &= "AND o.Situacao = 1" & vbCrLf
        End If

        If rbAberta.Checked Then strSQL &= "AND o.Estoque = 0" & vbCrLf

        If rbEncerrada.Checked Then strSQL &= "AND o.Estoque = 1" & vbCrLf

        If rbDtMov.Checked Then
            strSQL &= "AND o.Movimento between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If rbDtVen.Checked Then
            strSQL &= "AND o.Validade between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If rbOrdem.Checked Then strSQL &= "ORDER BY o.Ordem_Id" & vbCrLf

        If rbOPrd.Checked Then strSQL &= "ORDER BY OPXP.Produto_Id" & vbCrLf

        If rbOPrdNome.Checked Then strSQL &= "ORDER BY p.Nome" & vbCrLf

        Return strSQL

    End Function

    Private Function sqlReportPDF() As String

        Dim strSQL As String = "SELECT o.Empresa_Id, o.EndEmpresa_Id, EMP.Nome + ' - ' + EMP.Reduzido AS Empresa, o.Ordem_Id AS Ordem, OPXP.Produto_Id AS Produto, p.Nome AS ProdutoNome, OPXP.Lote, convert(varchar,o.Movimento,103) AS Movimento, convert(varchar,o.Validade,103) AS Validade, OPXP.Quantidade, OPXP.QuantidadeDeAjuste, " & vbCrLf &
                                                    "		case" & vbCrLf &
                                                    "			when o.Estoque = 0" & vbCrLf &
                                                    "				then 'NÃO'" & vbCrLf &
                                                    "				else 'SIM'" & vbCrLf &
                                                    "			end AS Estoque," & vbCrLf &
                                                    "		o.UsuarioInclusao, convert(varchar,o.UsuarioInclusaoData,103) AS UsuarioInclusaoData, " & vbCrLf &
                                                    "		isnull(o.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf &
                                                    "		case" & vbCrLf &
                                                    "			when isnull(o.UsuarioAlteracaoData,'') = ''" & vbCrLf &
                                                    "				then ''" & vbCrLf &
                                                    "				else convert(varchar,o.UsuarioAlteracaoData,103)" & vbCrLf &
                                                    "			end AS UsuarioAlteracaoData," & vbCrLf &
                                                    "		isnull(o.UsuarioCancelamento,'') AS UsuarioCancelamento, " & vbCrLf &
                                                    "		case" & vbCrLf &
                                                    "			when isnull(o.UsuarioCancelamentoData,'') = ''" & vbCrLf &
                                                    "				then ''" & vbCrLf &
                                                    "				else convert(varchar,o.UsuarioCancelamentoData,103) " & vbCrLf &
                                                    "			end AS UsuarioCancelamentoData, " & vbCrLf &
                                                    "			ROW_NUMBER() OVER (PARTITION BY  EMP.Reduzido, o.Ordem_Id ORDER BY EMP.Reduzido, o.Ordem_Id, opXL.Produto_Id) AS Sequencia, " & vbCrLf &
                                                    "			opXL.Produto_Id As ProdutoConsumo, " & vbCrLf &
                                                    "			pCons.Nome As ProdutoNomeConsumo, " & vbCrLf &
                                                    "			opXL.Lote_Id As LoteConsumo, " & vbCrLf &
                                                    "			opXL.Quantidade AS QuantidadeConsumo, " & vbCrLf &
                                                    "			CONVERT(VARCHAR,opXL.Validade,103) AS ValidadeConsumo " & vbCrLf &
                                                    "FROM OrdemDeProducao o " & vbCrLf &
                                                    "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                                                    "   ON o.Empresa_Id		            = OPXP.Empresa_Id " & vbCrLf &
                                                    "   AND o.EndEmpresa_Id	            = OPXP.EndEmpresa_Id " & vbCrLf &
                                                    "   AND o.Ordem_Id			        = OPXP.Ordem_Id " & vbCrLf &
                                                    "INNER JOIN Clientes As EMP " & vbCrLf &
                                                    "		ON o.Empresa_Id				= EMP.Cliente_Id " & vbCrLf &
                                                    "		AND o.EndEmpresa_Id			= EMP.Endereco_Id" & vbCrLf &
                                                    "INNER JOIN Produtos p" & vbCrLf &
                                                    "		ON p.Produto_id             = OPXP.Produto_Id" & vbCrLf &
                                                    "INNER JOIN OrdemDeProducaoXConsumo opXC " & vbCrLf &
                                                    "		ON o.Empresa_id             = opXC.Empresa_id" & vbCrLf &
                                                    "		AND o.EndEmpresa_Id         = opXC.EndEmpresa_Id" & vbCrLf &
                                                    "		AND o.Ordem_Id              = opXC.Ordem_Id" & vbCrLf &
                                                    "		AND OPXP.Produto_Id         = opXC.ProdutoProducao_Id" & vbCrLf &
                                                    "INNER JOIN OrdemDeProducaoXConsumoXLote opXL" & vbCrLf &
                                                    "		ON o.Empresa_id             = opXL.Empresa_id" & vbCrLf &
                                                    "		AND o.EndEmpresa_Id         = opXL.EndEmpresa_Id" & vbCrLf &
                                                    "		AND o.Ordem_Id              = opXL.Ordem_Id" & vbCrLf &
                                                    "		AND opXC.Produto_Id         = opXL.Produto_Id" & vbCrLf &
                                                    "INNER JOIN Produtos pCons" & vbCrLf &
                                                    "		ON opXL.Produto_Id = pCons.Produto_Id" & vbCrLf &
                                                    "Where o.Empresa_Id    = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                                    "  and o.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

        If rbCancelada.Checked Then
            strSQL &= "AND o.Situacao = 3" & vbCrLf
        Else
            strSQL &= "AND o.Situacao = 1" & vbCrLf
        End If

        If rbAberta.Checked Then strSQL &= "AND o.Estoque = 0" & vbCrLf

        If rbEncerrada.Checked Then strSQL &= "AND o.Estoque = 1" & vbCrLf

        If rbDtMov.Checked Then
            strSQL &= "AND o.Movimento between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If rbDtVen.Checked Then
            strSQL &= "AND o.Validade between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If rbOrdem.Checked Then strSQL &= "ORDER BY o.Ordem_Id" & vbCrLf

        If rbOPrd.Checked Then strSQL &= "ORDER BY OPXP.Produto_Id" & vbCrLf

        If rbOPrdNome.Checked Then strSQL &= "ORDER BY p.Nome" & vbCrLf


        Return strSQL

    End Function

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        HID.Value = Guid.NewGuid().ToString

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

End Class