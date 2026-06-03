Imports System.Drawing
Imports System.IO
Imports NGS.Lib.Negocio
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class RelatorioDoAlmoxarifado
    Inherits BasePage

    Private sql As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.ProducaoEstoque)

            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDoAlmoxarifado", "ACESSAR") Then
                    CargaUnidade()
                    Limpar()
                    HID.Value = Guid.NewGuid().ToString
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, " Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario"))
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        VerificaUnidade()
        txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.Text) Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData.Text) Then
            MsgBox(Me.Page, "Informe uma data válida.")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio()
        Try
            If Funcoes.VerificaPermissao("RelatorioDoAlmoxarifado", "RELATORIO") Then
                If Validar() Then

                    Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

                    Dim par As New ArrayList
                    If ucSelecaoProduto.TemSelecionado Then
                        par = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id")
                    End If

                    Dim sWhereProduto As String = ""
                    If ucSelecaoProduto.TemSelecionado Then
                        sWhereProduto = " AND " & par(0)
                    End If

                    sql = "SELECT PRODUTO, SUM(INICIAL) AS INICIAL, SUM(ENTRADA) AS ENTRADAS, SUM(SAIDA) AS SAIDAS, SUM(ENTRADA) - SUM(SAIDA) AS SALDO" & vbCrLf &
                            "INTO #ALMOXARIFADO" & vbCrLf &
                            "FROM (" & vbCrLf &
                            "	SELECT nXi.PRODUTO_ID AS PRODUTO, 0 AS INICIAL," & vbCrLf &
                            "	CASE " & vbCrLf &
                            "		WHEN nXi.ENTRADASAIDA_ID = 'E'" & vbCrLf &
                            "			THEN nXi.QUANTIDADEFISCAL" & vbCrLf &
                            "			ELSE 0" & vbCrLf &
                            "	END AS ENTRADA," & vbCrLf &
                            "	CASE " & vbCrLf &
                            "		WHEN nXi.ENTRADASAIDA_ID = 'S'" & vbCrLf &
                            "			THEN nXi.QUANTIDADEFISCAL" & vbCrLf &
                            "			ELSE 0" & vbCrLf &
                            "	END AS SAIDA" & vbCrLf &
                            "	FROM NOTASFISCAISxITENS nXi" & vbCrLf &
                            "	INNER JOIN NOTASFISCAIS n" & vbCrLf &
                            "	    ON n.empresa_id         = nXi.empresa_id" & vbCrLf &
                            "		AND n.endempresa_id     = nXi.endempresa_id" & vbCrLf &
                            "		AND n.cliente_id        = nXi.cliente_id" & vbCrLf &
                            "		AND n.endcliente_id     = nXi.endcliente_id" & vbCrLf &
                            "		AND n.entradasaida_id   = nXi.entradasaida_id" & vbCrLf &
                            "		AND n.serie_id          = nXi.serie_id" & vbCrLf &
                            "		AND n.nota_id           = nXi.nota_id" & vbCrLf &
                            "	INNER JOIN SubOperacoes SO" & vbCrLf &
                            "		ON nXi.Operacao         = SO.Operacao_Id" & vbCrLf &
                            "		AND nXi.SubOperacao     = SO.SubOperacoes_Id" & vbCrLf &
                            "	INNER JOIN PRODUTOS" & vbCrLf &
                            "		ON Produtos.PRODUTO_ID  = nXi.PRODUTO_ID" & vbCrLf &
                            "	WHERE nXi.EMPRESA_ID = '" & objEmpresa.Codigo & "'" & vbCrLf &
                            "	AND n.SITUACAO = 1" & vbCrLf &
                            "	AND n.MOVIMENTO <= '" & txtData.Text.ToSqlDate() & "'" & vbCrLf &
                            "	AND Produtos.ALMOXARIFADO = 1" & vbCrLf &
                            "	AND SO.EstoqueFiscal = 'S'" & vbCrLf &
                            "   " & sWhereProduto & vbCrLf &
                            "" & vbCrLf &
                            "	Union all" & vbCrLf &
                            "" & vbCrLf &
                            "	SELECT P.PRODUTO_ID AS PRODUTO," & vbCrLf &
                            "	CASE " & vbCrLf &
                            "		WHEN S.ESTOQUEINICIAL = 'S'" & vbCrLf &
                            "			THEN " & vbCrLf &
                            "				CASE " & vbCrLf &
                            "					WHEN S.ENTRADASAIDA = 'E'" & vbCrLf &
                            "						THEN P.ENTRADAS" & vbCrLf &
                            "						ELSE P.SAIDAS" & vbCrLf &
                            "				END" & vbCrLf &
                            "			ELSE 0" & vbCrLf &
                            "	END AS INICIAL," & vbCrLf &
                            "	CASE " & vbCrLf &
                            "		WHEN S.ENTRADASAIDA = 'E'" & vbCrLf &
                            "			THEN P.ENTRADAS" & vbCrLf &
                            "			ELSE 0" & vbCrLf &
                            "	END AS ENTRADA," & vbCrLf &
                            "	CASE " & vbCrLf &
                            "		WHEN S.ENTRADASAIDA = 'S'" & vbCrLf &
                            "			THEN P.SAIDAS" & vbCrLf &
                            "			ELSE 0" & vbCrLf &
                            "	END AS SAIDA" & vbCrLf &
                            "	FROM PRODUCAO P" & vbCrLf &
                            "		INNER JOIN SUBOPERACOES S" & vbCrLf &
                            "				ON S.OPERACAO_ID      = P.OPERACAO_ID" & vbCrLf &
                            "				AND S.SUBOPERACOES_ID = P.SUBOPERACAO_ID" & vbCrLf &
                            "		INNER JOIN PRODUTOS " & vbCrLf &
                            "				on Produtos.PRODUTO_ID     = P.PRODUTO_ID" & vbCrLf &
                            "	WHERE P.EMPRESA_ID = '" & objEmpresa.Codigo & "'" & vbCrLf &
                            "	AND P.MOVIMENTO_ID <= '" & txtData.Text.ToSqlDate() & "'" & vbCrLf &
                            "	AND Produtos.ALMOXARIFADO = 1" & vbCrLf &
                            "	AND P.FISICOFISCAL_ID = 2 " & vbCrLf &
                            "   " & sWhereProduto & vbCrLf &
                            "	) AS RESULTADO" & vbCrLf &
                            "GROUP BY RESULTADO.PRODUTO" & vbCrLf &
                            "" & vbCrLf &
                            "SELECT distinct Produto_id AS Produto, Data = MAX(Data_id) OVER (PARTITION BY Produto_id)" & vbCrLf &
                            "INTO #VERPRECO" & vbCrLf &
                            "FROM PRODUTOSxPRECOS" & vbCrLf &
                            "" & vbCrLf &
                            "SELECT A.PRODUTO, P.NOME, P.UNIDADE, A.INICIAL, A.ENTRADAS, A.SAIDAS, A.SALDO, ISNULL(PXP.VALOR,0) AS UNITARIO," & vbCrLf &
                            "CASE " & vbCrLf &
                            "	WHEN ISNULL(PXP.VALOR,0) = 0" & vbCrLf &
                            "		THEN 0" & vbCrLf &
                            "		ELSE A.SALDO * PXP.VALOR" & vbCrLf &
                            "END AS VALORATUAL" & vbCrLf &
                            "FROM #ALMOXARIFADO A" & vbCrLf &
                            "	INNER JOIN PRODUTOS P" & vbCrLf &
                            "			ON P.PRODUTO_ID = A.PRODUTO" & vbCrLf &
                            "	LEFT JOIN #VERPRECO VP" & vbCrLf &
                            "			ON VP.PRODUTO = A.PRODUTO" & vbCrLf &
                            "	LEFT JOIN PRODUTOSxPRECOS PxP" & vbCrLf &
                            "			ON PXP.PRODUTO_ID = VP.PRODUTO" & vbCrLf &
                            "			AND PXP.DATA_ID   = VP.DATA" & vbCrLf &
                            "WHERE P.Situacao = 1" & vbCrLf &
                            "ORDER BY P.NOME"

                    Dim ds As DataSet = Banco.ConsultaDataSet(sql, "almoxarifado")

                    If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Sem informação para o Período.", eTitulo.Info)
                        Exit Sub
                    End If

                    Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)

                            'criando planilha títulos
                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("POSIÇÃO DO ALMOXARIFADO")

                            'criando linha com o cabeçalho da planilha
                            Dim rowIndex As Integer = 1
                            Dim columnIndex As Integer = 1

                            'criando linha que informa o nome da empresa e o cnpj
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa a cidade e o estado da empresa
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa o título do relatório
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO DO ALMOXARIFADO")
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa o período selecionado na página
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO ATÉ : " & String.Format("{0:dd/MM/yyyy}", CDate(txtData.Text)))
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:I{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha com o cabeçalho da planilha
                            worksheet.Cells(rowIndex, columnIndex).Value = "Produto"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Nome"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Unidade"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Saldo Inicial"
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Entradas"
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Saídas"
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Saldo Atual"
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Valor Unitário"
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Valor Atual"
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            columnIndex += 1

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            rowIndex += 1

                            For Each row As DataRow In ds.Tables(0).Rows
                                columnIndex = 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("PRODUTO")
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("NOME")
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("UNIDADE")
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("INICIAL")
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("ENTRADAS")
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("SAIDAS")
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("SALDO")
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("UNITARIO")
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                columnIndex += 1

                                worksheet.Cells(rowIndex, columnIndex).Value = row("VALORATUAL")
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                columnIndex += 1

                                'aplicando formatação nas células do conteúdo
                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If

                                ''formatando células numéricas
                                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.000000_ ;[Red]-#,##0.000000"
                                worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                rowIndex += 1
                            Next

                            ''formatando células numéricas
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'criando colunas de totalizadores
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Value = String.Format("{0}", "VALOR TOTAL")

                            'criando colunas de totalizadores
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Formula = String.Format("=SUM(I6:I{0})", rowIndex - 1)

                            'setando autofit nas células da planilha
                            worksheet.Cells.AutoFitColumns(0)

                            'congelando quinta linha (cabeçalho)
                            worksheet.View.FreezePanes(6, 1)

                            package.Save()
                        End Using
                    End Using

                    Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class