Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports System.Drawing

Public Class BaseNova
    Inherits BasePage

    Private SqlArray As New ArrayList
    Private sql As String = String.Empty
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)

            If Not IsPostBack And IsConnect Then
                If Session("ssEmpresa") = "03189063000126" OrElse Session("ssEmpresa") = "05272759000147" Then
                    If Funcoes.VerificaPermissao("BaseNova", "ACESSAR") Then
                        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    Else
                        MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                        Exit Sub
                    End If
                Else
                    MsgBox(Me.Page, "Empresa sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a empresa.")
            ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                MsgBox(Me.Page, "Informe o período para consulta.")
            Else
                ListarNotas("C")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkGravar_Click(sender As Object, e As EventArgs) Handles lnkGravar.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a empresa.")
            ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                MsgBox(Me.Page, "Informe o período para gravação.")
            Else
                ListarNotas("I")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ListarNotas(ByVal tipo As String)
        Dim percentual As Decimal = 0
        Dim base As Decimal = 0

        Dim vlrIcms As Decimal = 0
        Dim vlrPis As Decimal = 0
        Dim vlrCofins As Decimal = 0

        Dim baseNova As Decimal = 0
        Dim pisNovo As Decimal = 0
        Dim cofinsNovo As Decimal = 0

        SqlArray.Clear()

        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        sql = "SELECT Empresa_Id, EndEmpresa_Id " & vbCrLf &
                 " FROM clientesXempresas " & vbCrLf

        sql &= "where left(Empresa_id,8) = '" & Left(ddlEmpresa.SelectedValue.ToString.Split("-")(0), 8) & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "clientesXempresas")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("DEMOSTRATIVO DE SAÍDA PIS/COFINS")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DEMOSTRATIVO DE SAÍDA PIS/COFINS")
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    worksheet.Cells(rowIndex, columnIndex).Value = "Empresa"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Cliente"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Nota"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Movimento"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Operação"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "CFOP"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Produto"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Base"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "%"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Icms"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Pis"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Cofins"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Base Nova"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Pis"
                    columnIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "Cofins"
                    columnIndex += 1
                    rowIndex += 1

                    'criando auto filtro na planilha
                    'worksheet.Cells("A3:G" & rowIndex).AutoFilter = True

                    For Each row As DataRow In ds.Tables(0).Rows

                        Dim listNotas = New [Lib].Negocio.ListNotasFiscais()

                        Dim listNotasSaida As New [Lib].Negocio.ListNotasFiscais(row("Empresa_Id"), row("EndEmpresa_Id"), CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "", "", "S", "")

                        For Each nf In listNotasSaida
                            If nf.CodigoOperacao = 21 Then
                                listNotas.Add(nf)
                            ElseIf nf.CodigoOperacao = 70 And nf.CodigoSubOperacao = 34 Then
                                listNotas.Add(nf)
                            End If
                        Next

                        Dim listNotasEntrada As New [Lib].Negocio.ListNotasFiscais(row("Empresa_Id"), row("EndEmpresa_Id"), CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), "", "", "E", "")

                        For Each nf In listNotasEntrada
                            If nf.CodigoOperacao = 1 And (nf.CodigoSubOperacao = 1 Or nf.CodigoSubOperacao = 9) Then
                                listNotas.Add(nf)
                            ElseIf nf.CodigoOperacao = 21 And (nf.CodigoSubOperacao = 4 Or nf.CodigoSubOperacao = 14) Then
                                listNotas.Add(nf)
                            End If
                        Next

                        'aplicando formatação nas células do cabeçalho
                        'Using range = worksheet.Cells(rowIndex, 1, rowIndex, listNotas.Where(Function(s) s.CodigoOperacao = 21).Count)
                        '    range.Style.Font.Bold = True
                        '    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        '    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        '    range.Style.Font.Color.SetColor(Color.White)
                        'End Using


                        For Each iNota As [Lib].Negocio.NotaFiscal In listNotas

                            For Each item In iNota.Itens
                                Dim temIcms As Boolean = item.Encargos.Where(Function(s) s.Codigo.Contains("ICMS")).Count > 0
                                Dim temPis As Boolean = item.Encargos.Where(Function(s) s.Codigo.Contains("PIS")).Count > 0

                                If temIcms And temPis Then

                                    If item.Encargos.Where(Function(s) s.Codigo.Contains("DESCONTOS")).Count > 0 Then
                                        base = item.Encargos.Where(Function(s) s.Codigo.Contains("LIQUIDO")).Sum(Function(s) s.Base)
                                    Else
                                        base = item.Encargos.Where(Function(s) s.Codigo.Contains("PRODUTO")).Sum(Function(s) s.Base)
                                    End If

                                    percentual = item.Encargos.Where(Function(s) s.Codigo.Contains("ICMS")).Sum(Function(s) s.Percentual)

                                    vlrIcms = item.Encargos.Where(Function(s) s.Codigo.Contains("ICMS")).Sum(Function(s) s.Valor)
                                    vlrPis = item.Encargos.Where(Function(s) s.Codigo.Contains("PIS")).Sum(Function(s) s.Valor)
                                    vlrCofins = item.Encargos.Where(Function(s) s.Codigo.Contains("COFINS")).Sum(Function(s) s.Valor)

                                    baseNova = 0
                                    pisNovo = 0
                                    cofinsNovo = 0

                                    columnIndex = 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = iNota.CodigoEmpresa
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = iNota.CodigoCliente
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = (iNota.Codigo.ToString + "-" + iNota.Serie)
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = iNota.Movimento
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = (iNota.CodigoOperacao.ToString + "-" + iNota.CodigoSubOperacao.ToString)
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = item.CFOP
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = item.CodigoProduto
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = base
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = percentual
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = vlrIcms
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = vlrPis
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = vlrCofins
                                    columnIndex += 1

                                    If percentual = 4 Then
                                        baseNova = Math.Round((base * CDec("0.8675")) / CDec("0.9075"), 2, MidpointRounding.AwayFromZero)
                                    ElseIf percentual = 7 Then
                                        baseNova = Math.Round((base * CDec("0.8375")) / CDec("0.9075"), 2, MidpointRounding.AwayFromZero)
                                    ElseIf percentual = 12 Then
                                        baseNova = Math.Round((base * CDec("0.7875")) / CDec("0.9075"), 2, MidpointRounding.AwayFromZero)
                                    ElseIf percentual = 17 Then
                                        baseNova = Math.Round((base * CDec("0.7375")) / CDec("0.9075"), 2, MidpointRounding.AwayFromZero)
                                    ElseIf percentual = 19.5 Then
                                        baseNova = Math.Round((base * CDec("0.7125")) / CDec("0.9075"), 2, MidpointRounding.AwayFromZero)
                                    End If

                                    If baseNova > 0 Then
                                        For Each enc In item.Encargos
                                            enc.BaseNova = baseNova

                                            sql = "Update NotasFiscaisXEncargos" & vbCrLf &
                                                    "   set BaseNova        = " & Str(enc.BaseNova) & vbCrLf

                                            If enc.Codigo.Contains("PIS") Then
                                                enc.ValorNovo = Math.Round((enc.BaseNova * enc.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                                                pisNovo = enc.ValorNovo

                                                sql &= "       ,ValorNovo      = " & Str(enc.ValorNovo) & vbCrLf
                                            ElseIf enc.Codigo.Contains("COFINS") Then
                                                enc.ValorNovo = Math.Round((enc.BaseNova * enc.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                                                cofinsNovo = enc.ValorNovo

                                                sql &= "       ,ValorNovo      = " & Str(enc.ValorNovo) & vbCrLf
                                            End If

                                            sql &= " Where  Empresa_Id     ='" & item.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                                    "   and EndEmpresa_Id   = " & item.NotaFiscal.EnderecoEmpresa & vbCrLf &
                                                    "   and Cliente_Id      ='" & item.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                                                    "   and EndCliente_Id   = " & item.NotaFiscal.EnderecoCliente & vbCrLf &
                                                    "   and EntradaSaida_Id ='" & item.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                                    "   and Nota_Id         = " & item.NotaFiscal.Codigo & vbCrLf &
                                                    "   and Serie_Id        ='" & item.NotaFiscal.Serie & "'" & vbCrLf &
                                                    "   and Produto_Id      ='" & item.CodigoProduto & "'" & vbCrLf &
                                                    "   and CFOP_ID         = " & item.CFOP & vbCrLf &
                                                    "   and Sequencia_id    = " & item.Sequencia & vbCrLf &
                                                    "   and Encargo_id      = '" & enc.Codigo & "'"

                                            SqlArray.Add(sql)
                                        Next
                                    End If

                                    worksheet.Cells(rowIndex, columnIndex).Value = baseNova
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = pisNovo
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = cofinsNovo
                                    columnIndex += 1

                                    'formatando células datas
                                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                                    ''formatando células numéricas
                                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                    'aplicando formatação nas células do conteúdo
                                    'If rowIndex Mod 2 = 0 Then
                                    '    Using range = worksheet.Cells(rowIndex, 1, rowIndex, listNotas.Count)
                                    '        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    '        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    '    End Using
                                    'End If

                                    rowIndex += 1
                                End If
                            Next
                        Next
                    Next

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    'worksheet.Cells(String.Format("G{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Formula = String.Format("=SUM(H6:H{0})", rowIndex - 1)
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    'worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    'worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
        End If

        If tipo = "I" Then
            If SqlArray IsNot Nothing AndAlso SqlArray.Count > 0 Then
                If Banco.GravaBanco(SqlArray) Then
                    Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

                    Contabilizar()
                Else
                    MsgBox(Me.Page, "Erro a Altererar Registro.")
                End If
            Else
                MsgBox(Me.Page, "Sem Registros para alteração.")
            End If
        Else
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        End If
    End Sub

    Private Sub Contabilizar()
        Try
            sql = "SELECT NotasFiscais.Empresa_Id," & vbCrLf &
                  "       NotasFiscais.EndEmpresa_Id," & vbCrLf &
                  "       NotasFiscais.Cliente_Id," & vbCrLf &
                  "       NotasFiscais.EndCliente_Id, " & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id," & vbCrLf &
                  "       NotasFiscais.Serie_Id," & vbCrLf &
                  "       NotasFiscais.Nota_Id" & vbCrLf &
                  "  FROM NotasFiscais" & vbCrLf &
                  " INNER JOIN SubOperacoes " & vbCrLf &
                  "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id " & vbCrLf &
                  "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                  " WHERE (NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "')" & vbCrLf &
                  "   AND (NotasFiscais.Situacao =  1)" & vbCrLf &
                  "   AND ((NotasFiscais.Operacao = 21 and entradasaida_id = 'S' ) OR" & vbCrLf &
                  "        (NotasFiscais.Operacao = 70 AND NotasFiscais.SubOperacao = 34) OR" & vbCrLf &
                  "        (NotasFiscais.Operacao =  1 AND (NotasFiscais.SubOperacao = 1  OR NotasFiscais.SubOperacao = 9)) OR" & vbCrLf &
                  " 	   (NotasFiscais.Operacao = 21 AND (NotasFiscais.SubOperacao = 4  OR NotasFiscais.SubOperacao = 14)))" & vbCrLf &
                  "   AND (SubOperacoes.Contabil = 'S')" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Notas")

            Dim Sqls As New ArrayList

            For Each row As DataRow In ds.Tables(0).Rows
                Dim nf As New [Lib].Negocio.NotaFiscal()
                nf.CodigoEmpresa = row("Empresa_Id")
                nf.EnderecoEmpresa = row("EndEmpresa_Id")
                nf.CodigoCliente = row("Cliente_Id")
                nf.EnderecoCliente = row("EndCliente_Id")
                nf.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nf.Codigo = row("Nota_Id")
                nf.Serie = row("Serie_Id")
                nf = New [Lib].Negocio.NotaFiscal(nf)

                nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                nf.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                nf.Razao.ContabilizarNotaSql(Sqls, 10)
            Next

            If Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, "Processo realizado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Info, True)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info, True)
        End Try
    End Sub
End Class