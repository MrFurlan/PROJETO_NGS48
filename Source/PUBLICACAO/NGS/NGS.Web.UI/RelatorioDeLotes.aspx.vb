Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports System.Drawing

Public Class RelatorioDeLotes
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioDeLotes", "ACESSAR") Then
                CarregarUnidade()
                Limpar()
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        pnlDataMovimento.Visible = False
        tipoLista.Visible = False
        HID.Value = Guid.NewGuid().ToString

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub chkPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPeriodo.CheckedChanged
        pnlDataMovimento.Visible = chkPeriodo.Checked
        tipoLista.Visible = chkPeriodo.Checked
    End Sub

    Function Validar() As Boolean
        If ddlUnidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Unidade de Negocio")
            Return False
        ElseIf DdlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Empresa")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function VerificarGrupoProduto(ByRef Sql As String, ByVal nPrd As String) As String
        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", nPrd)
            Sql &= "  and " & retorno(0) & vbCrLf
            Return retorno(1)
        Else
            Return ""
        End If
    End Function

    Function getDataSet(ByVal objEmpresa As [Lib].Negocio.Cliente) As DataSet

        'Sql = "SELECT nXt.Produto_Id AS Produto, nXt.Lote_Id As Lote, nXt.Fabricado, nXt.Validade," & vbCrLf & _
        '        "    sum(Case" & vbCrLf & _
        '        "            when nXt.EntradaSaida_Id = 'E'" & vbCrLf & _
        '        "            then nXt.Quantidade - isnull(nXt.QuantidadeDeConsumo,0)" & vbCrLf & _
        '        "            else nXt.Quantidade * -1 " & vbCrLf & _
        '        "        end) Quantidade" & vbCrLf & _
        '        "INTO #Consumo" & vbCrLf & _
        '        "FROM NotaFiscalXLote nXt" & vbCrLf & _
        '        "	inner join NotasFiscais n" & vbCrLf & _
        '        "		    ON n.Empresa_Id       = nXt.Empresa_Id" & vbCrLf & _
        '        "		    and n.EndEmpresa_iD   = nXt.EndEmpresa_Id" & vbCrLf & _
        '        "			and n.Cliente_Id      = nXt.Cliente_Id" & vbCrLf & _
        '        "			and n.EndCliente_Id   = nXt.EndCliente_Id" & vbCrLf & _
        '        "			and n.EntradaSaida_Id = nXt.EntradaSaida_Id" & vbCrLf & _
        '        "			and n.Serie_Id        = nXt.Serie_Id" & vbCrLf & _
        '        "			and n.Nota_Id         = nXt.Nota_Id" & vbCrLf & _
        '        "WHERE nXt.Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and nXt.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf & _
        '        "  and n.Situacao        = 1" & vbCrLf

        'VerificarGrupoProduto(Sql, "nXt.Produto_id")

        'Sql &= "GROUP BY nXt.Produto_Id, nXt.Lote_Id, nXt.Fabricado, nXt.Validade" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "Insert into #Consumo (Produto, Lote, Fabricado, Validade, Quantidade)" & vbCrLf & _
        '        "select Produto, Lote, Movimento AS Fabricado, Validade," & vbCrLf & _
        '        "		sum(case" & vbCrLf & _
        '        "			when QuantidadeDeAjuste > 0" & vbCrLf & _
        '        "				then QuantidadeDeAjuste" & vbCrLf & _
        '        "				else Quantidade" & vbCrLf & _
        '        "			end) AS Quantidade" & vbCrLf & _
        '        "from OrdemDeProducao" & vbCrLf & _
        '        "WHERE Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf & _
        '        "  and Situacao      = 1" & vbCrLf

        'VerificarGrupoProduto(Sql, "Produto")

        'Sql &= "group by Produto, Lote, Movimento, Validade" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "Insert into #Consumo (Produto, Lote, Fabricado, Validade, Quantidade)" & vbCrLf & _
        '        "SELECT Produto_Id AS Produto, Lote_Id As Lote, Movimento AS Fabricado, Validade, Quantidade" & vbCrLf & _
        '        "FROM OrdemDeProducaoInterna" & vbCrLf & _
        '        "WHERE Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf & _
        '        "  and EntradaSaida  = 'E'" & vbCrLf

        'VerificarGrupoProduto(Sql, "Produto_id")

        'Sql &= "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "SELECT Empresa_Id, EndEmpresa_Id, Produto, Lote, Movimento As Fabricado" & vbCrLf & _
        '        "INTO #OrdemPRD" & vbCrLf & _
        '        "FROM OrdemDeProducao" & vbCrLf & _
        '        "WHERE Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf & _
        '        "  and Situacao = 1" & vbCrLf

        'VerificarGrupoProduto(Sql, "Produto")

        'Sql &= "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "INSERT INTO #OrdemPRD (Empresa_Id, EndEmpresa_Id, Produto, Lote, Fabricado)" & vbCrLf & _
        '        "SELECT Empresa_Id, EndEmpresa_Id, Produto_Id, Lote_Id, Movimento FROM OrdemDeProducaoInterna" & vbCrLf & _
        '        "WHERE Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf

        'VerificarGrupoProduto(Sql, "Produto_id")

        'Sql &= "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "SELECT opXL.Produto_Id AS Produto, opXL.Lote_Id As Lote," & vbCrLf & _
        '        "		case" & vbCrLf & _
        '        "			when SUBSTRING(opXL.Lote_Id,4,1) = '.'" & vbCrLf & _
        '        "				then odP.Fabricado" & vbCrLf & _
        '        "				else nXt.Fabricado" & vbCrLf & _
        '        "			end Fabricado," & vbCrLf & _
        '        "opXL.Validade, SUM(opXL.Quantidade) AS Quantidade" & vbCrLf & _
        '        "INTO #ConsumoLote" & vbCrLf & _
        '        "FROM OrdemDeProducaoXConsumoXLote opXL" & vbCrLf & _
        '        "	INNER JOIN OrdemDeProducao op" & vbCrLf & _
        '        "			on op.Empresa_id     = opXL.Empresa_id" & vbCrLf & _
        '        "			and op.EndEmpresa_Id = opXL.EndEmpresa_Id" & vbCrLf & _
        '        "			and op.Ordem_Id      = opXL.Ordem_Id" & vbCrLf & _
        '        "	LEFT JOIN (select Empresa_id, EndEmpresa_Id, Produto_Id, Lote_Id, Max(Fabricado) AS Fabricado " & vbCrLf & _
        '        "				from NotaFiscalXLote" & vbCrLf & _
        '        "				group by Empresa_id, EndEmpresa_Id, Produto_Id, Lote_Id) AS nXt" & vbCrLf & _
        '        "			on nXt.Empresa_Id     = opXL.Empresa_id" & vbCrLf & _
        '        "			and nXt.EndEmpresa_Id = opXL.EndEmpresa_Id" & vbCrLf & _
        '        "			and nXt.Produto_Id    = opXL.Produto_Id" & vbCrLf & _
        '        "			and nXt.Lote_Id       = opXL.Lote_Id" & vbCrLf & _
        '        "	LEFT JOIN (select Empresa_id, EndEmpresa_Id, Produto, Lote, Max(Fabricado) AS Fabricado" & vbCrLf & _
        '        "				from #OrdemPRD" & vbCrLf & _
        '        "				group by Empresa_id, EndEmpresa_Id, Produto, Lote) AS odP" & vbCrLf & _
        '        "			on odP.Empresa_Id     = opXL.Empresa_id" & vbCrLf & _
        '        "			and odP.EndEmpresa_Id = opXL.EndEmpresa_Id" & vbCrLf & _
        '        "			and odp.Produto       = opXL.Produto_Id" & vbCrLf & _
        '        "			and odP.Lote          = opXL.Lote_Id" & vbCrLf & _
        '        "WHERE opXL.Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and opXL.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf & _
        '        "  and op.Situacao        = 1" & vbCrLf

        'VerificarGrupoProduto(Sql, "opXL.Produto_id")

        'Sql &= "GROUP BY opXL.Produto_Id, opXL.Lote_Id, nXt.Fabricado, odP.Fabricado, opXL.Validade" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "SELECT Produto_Id AS Produto, Lote_Id As Lote, Movimento AS Fabricado, Validade, Quantidade" & vbCrLf & _
        '        "INTO #ConsumoLoteInterno" & vbCrLf & _
        '        "FROM OrdemDeProducaoInterna" & vbCrLf & _
        '        "WHERE Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf & _
        '        "  and EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf & _
        '        "  and EntradaSaida  = 'S'" & vbCrLf

        'VerificarGrupoProduto(Sql, "Produto_id")

        'Sql &= "" & vbCrLf & _
        '        "" & vbCrLf & _
        '        "SELECT c.Produto, p.Nome, c.Lote, convert(varchar,c.Fabricado,103) as Fabricado, convert(varchar,c.Validade,103) as Validade, " & vbCrLf & _
        '        "		sum(c.Quantidade - isnull(cl.Quantidade,0) - isnull(clI.Quantidade,0)) as Quantidade" & vbCrLf & _
        '        "FROM #Consumo c" & vbCrLf & _
        '        "	INNER JOIN Produtos p" & vbCrLf & _
        '        "			on p.Produto_Id = c.Produto" & vbCrLf & _
        '        "	LEFT JOIN #ConsumoLote cl" & vbCrLf & _
        '        "			on cl.Produto    = c.Produto" & vbCrLf & _
        '        "			and cl.Lote      = c.Lote" & vbCrLf & _
        '        "			and cl.Fabricado = c.Fabricado" & vbCrLf & _
        '        "			and cl.Validade  = c.Validade" & vbCrLf & _
        '        "	LEFT JOIN #ConsumoLoteInterno clI" & vbCrLf & _
        '        "			on clI.Produto    = c.Produto" & vbCrLf & _
        '        "			and clI.Lote      = c.Lote" & vbCrLf & _
        '        "			and clI.Fabricado = c.Fabricado" & vbCrLf & _
        '        "			and clI.Validade  = c.Validade" & vbCrLf

        'If chkPeriodo.Checked Then
        '    If rdMovimento.Checked Then
        '        Sql &= " WHERE c.Fabricado BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf
        '    Else
        '        Sql &= " WHERE c.Validade BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf
        '    End If
        'End If

        'Sql &= "GROUP BY c.Produto, p.Nome, c.Lote, c.Fabricado, c.Validade" & vbCrLf & _
        '        "HAVING sum(c.Quantidade - isnull(cl.Quantidade,0) - isnull(clI.Quantidade,0)) > 0" & vbCrLf & _
        '        "ORDER BY p.Nome, c.Validade"


        Sql = String.Empty
        Dim pGRupo As Boolean = False

        ' Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "RelatorioDeLotes")

        VerificarGrupoProduto(Sql, "proD")

        If Sql.Contains("LEFT(grupo") Then
            Sql = Sql.Replace("and proD in (Select Produto_id from Produtos where LEFT(grupo,5) in (", "")
            Sql = Sql.Replace("))", "")
            Sql = Trim(Sql)
            pGRupo = True
        Else
            Sql = Sql.Replace("and proD in (Select Produto_id from Produtos where Produto_Id in (", "")
            Sql = Sql.Replace(")", "")
            Sql = Trim(Sql)
        End If

        Dim pLote = New OrdemParaProducaoXConsumo()

        Dim ds As New DataSet
        ds = pLote.buscarLoteDeFornecedor(objEmpresa.Codigo, objEmpresa.CodigoEndereco, Sql, pGRupo)

        Return ds
    End Function

    Private Sub GerarExcel()
        Try

            Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

            Dim ds As DataSet = getDataSet(objEmpresa)

            Dim dtLoteFornecedor As New DataTable("ItemLoteFornecedor")
            dtLoteFornecedor.Columns.Add("Produto", Type.GetType("System.String"))
            dtLoteFornecedor.Columns.Add("Nome", Type.GetType("System.String"))
            dtLoteFornecedor.Columns.Add("Lote", Type.GetType("System.String"))
            dtLoteFornecedor.Columns.Add("Fabricado", Type.GetType("System.String"))
            dtLoteFornecedor.Columns.Add("Validade", Type.GetType("System.String"))
            dtLoteFornecedor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))


            Dim dt As DataTable = New DataTable()
            dt = ds.Tables(0)

            Dim ds1 As DataView
            ds1 = New DataView(dt)

            ds1.Sort = "Produto ASC"

            For Each row As DataRow In ds1.ToTable.Rows

                Dim prd As Produto = New Produto(row("Produto"))

                Dim newRow As DataRow = dtLoteFornecedor.NewRow()

                newRow("Produto") = row("Produto")
                newRow("Nome") = prd.Nome
                newRow("Lote") = row("Lote")
                newRow("Fabricado") = CDate(row("Fabricado")).ToString("dd/MM/yyyy")
                newRow("Validade") = CDate(row("Validade")).ToString("dd/MM/yyyy")
                newRow("Quantidade") = row("Quantidade")

                dtLoteFornecedor.Rows.Add(newRow)

            Next

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("RELATÓRIO DE LOTES")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELATÓRIO DE LOTES")
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    If chkPeriodo.Checked Then
                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                        If rdMovimento.Checked Then
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "MOVIMENTO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                        Else
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "VALIDADE : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                        End If

                        worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1
                    End If

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dtLoteFornecedor.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    If chkPeriodo.Checked Then
                        worksheet.Cells("C5:E" & rowIndex).AutoFilter = True
                    Else
                        worksheet.Cells("C4:E" & rowIndex).AutoFilter = True
                    End If

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dtLoteFornecedor.Rows
                        columnIndex = 1
                        For Each col As DataColumn In dtLoteFornecedor.Columns

                            If col.ColumnName = "Quantidade" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToDecimal(row(col.ColumnName))
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If

                            columnIndex += 1
                        Next

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dtLoteFornecedor.Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'formatando células numéricas
                    If chkPeriodo.Checked Then
                        worksheet.Cells(String.Format("F6:F{0}", rowIndex)).Style.Numberformat.Format = "0.0000_ ;[Red]-0.0000"
                    Else
                        worksheet.Cells(String.Format("F5:F{0}", rowIndex)).Style.Numberformat.Format = "0.0000_ ;[Red]-0.0000"
                    End If

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    If chkPeriodo.Checked Then
                        worksheet.View.FreezePanes(6, 1)
                    Else
                        worksheet.View.FreezePanes(5, 1)
                    End If

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioDeLotes", "RELATORIO") Then
                If Validar() Then GerarExcel()
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir o relatorio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeLotes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class