Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Drawing
Public Class RelatorioDePagamentos
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioDePagamentos", "ACESSAR") Then
                CarregarUnidade()
                Limpar()
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Financeiro.aspx")
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

    Function Validar() As Boolean
        If ddlUnidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Unidade de Negocio")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteInfLT" & HID.Value, "txtNome")
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            If Validar() Then EmitirRelatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub EmitirRelatorio()
        Try
            If Funcoes.VerificaPermissao("RelatorioDePagamentos", "RELATORIO") Then
                Dim Periodo As String = "PERÍODO: " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text))

                Dim objEmpresa As New Cliente()

                If DdlEmpresa.SelectedIndex > 0 Then
                    objEmpresa = New Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))
                Else
                    objEmpresa = New Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
                End If

                Dim ds As DataSet = getDataSet(objEmpresa)

                If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Sem informações para os parâmetros selecionados.", eTitulo.Info)
                    Exit Sub
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

                        worksheet.Cells("A1:B1").Merge = True
                        worksheet.Cells("A2:B2").Merge = True
                        worksheet.Cells("A3:B3").Merge = True
                        worksheet.Cells("A4:B4").Merge = True
                        worksheet.Cells("A5:B5").Merge = True
                        worksheet.Cells("A6:B6").Merge = True
                        worksheet.Cells("A7:B7").Merge = True
                        worksheet.Cells("A1:B7").Style.Font.Bold = True
                        worksheet.Cells("A1:B7").Style.Font.Size = 14

                        worksheet.Cells("A1").Value = objEmpresa.Nome
                        worksheet.Cells("A2").Value = "           " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                        worksheet.Cells("A3").Value = "           CNPJ - " & objEmpresa.CodigoFormatado

                        'criando linha que informa o título do relatório
                        worksheet.Cells("A5").Style.Font.Bold = True
                        worksheet.Cells("A5").Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells("A5").Value = "           RESUMO DOS PAGAMENTOS"

                        worksheet.Cells("A7").Value = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables("Pagamentos").Columns
                            If col.ColumnName = "Tipo" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = "DESCRIÇÃO"
                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Valor" Then
                                worksheet.Cells(rowIndex, 2).Value = "R$ VALOR"
                                worksheet.Cells(rowIndex, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                worksheet.Cells(rowIndex, 2).Style.Font.Size = 12
                            End If

                            columnIndex += 1
                        Next

                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("Pagamentos").Columns.Count - 1)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        rowIndex += 1

                        Dim descricao As String = String.Empty
                        Dim grupo As String = String.Empty

                        'criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In ds.Tables("Pagamentos").Rows
                            columnIndex = 1

                            For Each col As DataColumn In ds.Tables("Pagamentos").Columns
                                If columnIndex = 1 Then
                                    If Not descricao = row(col.ColumnName) Then

                                        If grupo.Length > 0 Then
                                            rowIndex += 2
                                        End If

                                        descricao = row(col.ColumnName)

                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                        worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                                        worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 10
                                    End If
                                ElseIf columnIndex = 2 Then

                                    grupo = row(col.ColumnName)

                                    If grupo.Length > 0 Then
                                        rowIndex += 1
                                        worksheet.Cells(rowIndex, 1).Value = row(col.ColumnName)
                                        worksheet.Cells(rowIndex, 1).Style.Font.Size = 10
                                    End If
                                Else
                                    worksheet.Cells(rowIndex, 2).Value = row(col.ColumnName)
                                    worksheet.Cells(rowIndex, 2).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(rowIndex, 2).Style.Font.Size = 10
                                End If

                                columnIndex += 1
                            Next

                            If grupo.Length = 0 Then
                                rowIndex += 2
                            End If
                        Next

                        worksheet.Cells(rowIndex, 1).Value = "TOTAL"
                        worksheet.Cells(rowIndex, 1).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(rowIndex, 1).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, 1).Style.Font.Size = 10

                        worksheet.Cells(String.Format("B{0}", rowIndex)).Formula = String.Format("=SUM(B6:B{0})", rowIndex - 1)
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Font.Size = 10

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando quinta linha (cabeçalho)
                        worksheet.View.FreezePanes(9, 1)

                        '******************************************
                        '*** 2 - 'Criando Lista de Títulos
                        '******************************************
                        worksheet = package.Workbook.Worksheets.Add("Titulos")
                        worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                        'criando linha com o cabeçalho da planilha
                        rowIndex = 1
                        columnIndex = 1

                        'criando linha que informa o nome da empresa e o cnpj
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa a cidade e o estado da empresa
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o título do relatório
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELAÇÃO DOS PAGAMENTOS")
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables(1).Columns
                            If col.ColumnName = "Tipo" Then
                                worksheet.Cells(rowIndex, 1).Value = "DESCRIÇÃO"
                                worksheet.Cells(rowIndex, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 1).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Cidade" Then
                                worksheet.Cells(rowIndex, 2).Value = "EMPRESA"
                                worksheet.Cells(rowIndex, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 2).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Titulo" Then
                                worksheet.Cells(rowIndex, 3).Value = "TÍTULO"
                                worksheet.Cells(rowIndex, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 3).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Prorrogacao" Then
                                worksheet.Cells(rowIndex, 4).Value = "VENCIMENTO"
                                worksheet.Cells(rowIndex, 4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 4).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Cliente" Then
                                worksheet.Cells(rowIndex, 5).Value = "CLIENTE"
                                worksheet.Cells(rowIndex, 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 5).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Nome" Then
                                worksheet.Cells(rowIndex, 6).Value = "NOME"
                                worksheet.Cells(rowIndex, 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 6).Style.Font.Size = 12
                            ElseIf col.ColumnName = "Historico" Then
                                worksheet.Cells(rowIndex, 8).Value = "HISTÓRICO"
                                worksheet.Cells(rowIndex, 8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 8).Style.Font.Size = 12
                            ElseIf col.ColumnName = "ValorLiquido" Then
                                worksheet.Cells(rowIndex, 7).Value = "R$ VALOR"
                                worksheet.Cells(rowIndex, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                worksheet.Cells(rowIndex, 7).Style.Font.Size = 12
                            End If
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:F" & rowIndex).AutoFilter = True

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count - 3)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using
                        rowIndex += 1

                        descricao = String.Empty
                        Dim primeiraVez As Boolean = True

                        For Each row As DataRow In ds.Tables(1).Rows
                            columnIndex = 1

                            If Not descricao = row.Item(0) Then
                                If primeiraVez Then
                                    primeiraVez = False
                                Else
                                    rowIndex += 2
                                End If

                                descricao = row.Item(0)

                                worksheet.Cells(rowIndex, columnIndex).Value = row.Item(0)
                                worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                                worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 10

                                rowIndex += 1
                            End If

                            'Grupo
                            worksheet.Cells(rowIndex, 1).Value = row.Item(1)
                            'Cidade
                            worksheet.Cells(rowIndex, 2).Value = row.Item(6)
                            'Título
                            worksheet.Cells(rowIndex, 3).Value = row.Item(2)
                            'Vencimento
                            worksheet.Cells(rowIndex, 4).Value = row.Item(3)
                            'Cliente
                            worksheet.Cells(rowIndex, 5).Value = row.Item(7) & "-" & row.Item(8)
                            'Nome
                            worksheet.Cells(rowIndex, 6).Value = row.Item(9)
                            'Histórico
                            worksheet.Cells(rowIndex, 7).Value = row.Item(11)
                            'Valor
                            worksheet.Cells(rowIndex, 8).Value = row.Item(10)

                            'formatando células datas
                            worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                            'formatando células numéricas
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            rowIndex += 1
                        Next

                        rowIndex += 1

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True


                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Formula = String.Format("=SUM(G6:G{0})", rowIndex - 1)
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando quinta linha (cabeçalho)
                        worksheet.View.FreezePanes(6, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                'download do arquivo pelo browser
                'Download(fileName)
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir o relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataSet(ByVal objEmpresa As [Lib].Negocio.Cliente) As DataSet

        Sql = "SELECT tit.Registro_id as Titulo, tit.Prorrogacao, tit.Empresa, tit.EndEmpresa, tit.Cliente, tit.EndCliente, " & vbCrLf & _
                "		n.EntradaSaida_id as EntradaSaida, n.Nota_Id as Nota, n.Serie_id as Serie, " & vbCrLf & _
                "		cXp.ContaClientes, cXp.Descricao AS DescricaoCarteira, cXtri.Tributo_Id AS DescricaoTributo," & vbCrLf & _
                "		plCarteira.Adiantamento, so.Operacao_Id, so.GrupoDeContas," & vbCrLf & _
                "		case" & vbCrLf & _
                "			when isnull(tit.NumeroDoCheque,0) > 0" & vbCrLf & _
                "				then tit.Historico + ' - CHEQUE ' + CONVERT(VARCHAR,tit.NumeroDoCheque)" & vbCrLf & _
                "				else tit.Historico" & vbCrLf & _
                "		end as Historico," & vbCrLf & _
                "		tit.ValorLiquido" & vbCrLf & _
                "INTO #Teste" & vbCrLf & _
                "FROM ContasAPagar tit" & vbCrLf & _
                "	LEFT JOIN ComprasXProdutos AS cXp " & vbCrLf & _
                "			ON cXp.Produto_Id   = tit.Carteira " & vbCrLf & _
                "	LEFT JOIN CarteirasXTributos cXtri" & vbCrLf & _
                "			ON cXtri.Carteira_Id  = tit.Carteira" & vbCrLf & _
                "			AND cXtri.Tributo_Id  = tit.Tributo" & vbCrLf & _
                "	LEFT JOIN PlanoDeContas AS plCarteira " & vbCrLf & _
                "			ON plCarteira.Conta_Id   = cXp.ContaClientes " & vbCrLf & _
                "	LEFT JOIN NotaFiscalXTitulo AS nXt " & vbCrLf & _
                "			ON nXt.Empresa_Id     = tit.Empresa" & vbCrLf & _
                "			AND nXt.EndEmpresa_Id = tit.EndEmpresa" & vbCrLf & _
                "			AND nXt.Cliente_Id    = tit.Cliente" & vbCrLf & _
                "			AND nXt.EndCliente_Id = tit.EndCliente" & vbCrLf & _
                "			AND nXt.Titulo_Id     = tit.Registro_Id" & vbCrLf & _
                "	LEFT JOIN NotasFiscais n" & vbCrLf & _
                "			ON n.Empresa_id       = nXt.Empresa_id" & vbCrLf & _
                "			AND n.EndEmpresa_id   = nXt.EndEmpresa_id" & vbCrLf & _
                "			AND n.Cliente_id      = nXt.Cliente_id" & vbCrLf & _
                "			AND n.EndCliente_id   = nXt.EndCliente_id" & vbCrLf & _
                "			AND n.EntradaSaida_id = nXt.EntradaSaida_id" & vbCrLf & _
                "			AND n.Serie_id        = nXt.Serie_id" & vbCrLf & _
                "			AND n.Nota_id         = nXt.Nota_id" & vbCrLf & _
                "	LEFT JOIN SubOperacoes so" & vbCrLf & _
                "			ON so.Operacao_Id      = n.Operacao" & vbCrLf & _
                "			AND so.SubOperacoes_Id = n.SubOperacao" & vbCrLf & _
                "Where 1 = 1 " & vbCrLf & _
                "AND (tit.Situacao IN(1,101,102)) " & vbCrLf & _
                "AND (tit.Grupado <> 'M') " & vbCrLf & _
                "AND tit.UnidadeDeNegocio = '" & ddlUnidade.SelectedValue & "'" & vbCrLf

        If DdlEmpresa.SelectedIndex > 0 Then Sql &= "AND tit.Empresa = '" & objEmpresa.Codigo & "'" & vbCrLf

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= "AND tit.Cliente = '" & strCliente(0) & "' And tit.EndCliente = " & strCliente(1) & vbCrLf
        End If

        If radVencimento.Checked Then
            Sql &= "AND tit.Prorrogacao BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
        Else
            Sql &= "AND tit.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
        End If

        If Left(Session("ssEmpresa"), 8) = "03189063" Then
            Sql &= "AND ((tit.Provisao = 2) OR (tit.Provisao = 1 AND tit.TipoPagto = 2))" & vbCrLf
            Sql &= "AND isnull(nXt.Cliente_Id,'') not in ('05841957000184')" & vbCrLf
        Else
            Sql &= "AND (tit.Provisao = 2)" & vbCrLf
        End If

        Sql &= "" & vbCrLf & _
                "" & vbCrLf & _
                "SELECT #Teste.Empresa, #Teste.EndEmpresa, #Teste.Cliente, #Teste.EndCliente, " & vbCrLf & _
                "		#Teste.EntradaSaida, #Teste.Serie, #Teste.Nota, ni.OperacaoXEstado " & vbCrLf & _
                "into #ToperacaXEstado" & vbCrLf & _
                "FROM NotasFiscaisXItens ni" & vbCrLf & _
                "	INNER JOIN #Teste " & vbCrLf & _
                "			ON ni.Empresa_id       = #Teste.Empresa" & vbCrLf & _
                "			AND ni.EndEmpresa_id   = #Teste.EndEmpresa" & vbCrLf & _
                "			AND ni.Cliente_id      = #Teste.Cliente" & vbCrLf & _
                "			AND ni.EndCliente_id   = #Teste.EndCliente" & vbCrLf & _
                "			AND ni.EntradaSaida_id = #Teste.EntradaSaida" & vbCrLf & _
                "			AND ni.Serie_id        = #Teste.Serie" & vbCrLf & _
                "			AND ni.Nota_id         = #Teste.Nota" & vbCrLf & _
                "			AND ni.Sequencia_Id    = 1" & vbCrLf & _
                "GROUP BY #Teste.Empresa, #Teste.EndEmpresa, #Teste.Cliente, #Teste.EndCliente, #Teste.EntradaSaida, #Teste.Serie, #Teste.Nota, ni.OperacaoXEstado" & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "select " & vbCrLf & _
                "	case " & vbCrLf & _
                "		when t.Adiantamento = 1" & vbCrLf & _
                "		then pla.Titulo" & vbCrLf & _
                "		else" & vbCrLf & _
                "			case" & vbCrLf & _
                "				when t.Nota is NULL" & vbCrLf & _
                "				then 'OUTRAS DESPESAS'" & vbCrLf & _
                "				else " & vbCrLf & _
                "					case" & vbCrLf & _
                "						when t.Operacao_id in(1,2)" & vbCrLf & _
                "						then plaOpe.Titulo" & vbCrLf & _
                "						else plaOut.Titulo" & vbCrLf & _
                "					end" & vbCrLf & _
                "			end" & vbCrLf & _
                "	end as Tipo, " & vbCrLf & _
                "	case" & vbCrLf & _
                "		when t.Operacao_id in(1,2) " & vbCrLf & _
                "		then ge.Descricao" & vbCrLf & _
                "		else " & vbCrLf & _
                "			case" & vbCrLf & _
                "				when t.Nota is NULL" & vbCrLf & _
                "				then t.DescricaoTributo" & vbCrLf & _
                "				else ''" & vbCrLf & _
                "			end" & vbCrLf & _
                "	end as Grupo," & vbCrLf & _
                "	t.Titulo, t.Prorrogacao, t.Empresa, t.EndEmpresa, t.Cliente, t.EndCliente, cli.Nome, t.Historico, t.ValorLiquido," & vbCrLf & _
                "	case" & vbCrLf & _
                "		when t.Operacao_id in(1,2)" & vbCrLf & _
                "		then 1" & vbCrLf & _
                "		else 2" & vbCrLf & _
                "	end AS listaGrupo" & vbCrLf & _
                "into #Rel1" & vbCrLf & _
                "from #Teste t" & vbCrLf & _
                "	INNER JOIN Clientes cli" & vbCrLf & _
                "			ON cli.Cliente_Id   = t.Cliente" & vbCrLf & _
                "			AND cli.Endereco_Id = t.EndCliente" & vbCrLf & _
                "	LEFT JOIN #ToperacaXEstado cod" & vbCrLf & _
                "			ON cod.Empresa       = t.Empresa" & vbCrLf & _
                "			AND cod.EndEmpresa   = t.EndEmpresa" & vbCrLf & _
                "			AND cod.Cliente      = t.Cliente" & vbCrLf & _
                "			AND cod.EndCliente   = t.EndCliente" & vbCrLf & _
                "			AND cod.EntradaSaida = t.EntradaSaida" & vbCrLf & _
                "			AND cod.Serie        = t.Serie" & vbCrLf & _
                "			AND cod.Nota         = t.Nota" & vbCrLf & _
                "	LEFT JOIN OperacaoXEstado oXe" & vbCrLf & _
                "			ON oXe.Codigo_Id     = cod.OperacaoXEstado" & vbCrLf & _
                "	LEFT JOIN OperacaoXEstadoXEncargo oXeXe" & vbCrLf & _
                "			ON oXeXe.Codigo_Id   = oXe.Codigo_Id" & vbCrLf & _
                "			AND oXeXe.Encargo_Id = 'PRODUTO'" & vbCrLf & _
                "	LEFT JOIN PlanoDeContas AS pla" & vbCrLf & _
                "			ON pla.Conta_Id      = t.ContaClientes" & vbCrLf & _
                "	LEFT JOIN PlanoDeContas AS plaOpe" & vbCrLf & _
                "			ON plaOpe.Conta_Id   = t.GrupoDeContas" & vbCrLf & _
                "	LEFT JOIN PlanoDeContas AS plaOut" & vbCrLf & _
                "			ON plaOut.Conta_Id   = left(oXeXe.DebitaConta,7)" & vbCrLf & _
                "	LEFT JOIN GruposDeEstoques ge" & vbCrLf & _
                "			ON	ge.Grupo_id      = oXe.GrupoProduto" & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "SELECT Tipo, " & vbCrLf & _
                "		case" & vbCrLf & _
                "			when listaGrupo = 1" & vbCrLf & _
                "				then Grupo" & vbCrLf & _
                "				else ''" & vbCrLf & _
                "		end	AS Grupo," & vbCrLf & _
                "		Titulo, Empresa, EndEmpresa, Cliente, EndCliente, Nome, Historico, ValorLiquido " & vbCrLf & _
                "into #Rel2" & vbCrLf & _
                "FROM #Rel1" & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "SELECT Tipo, Grupo, Sum(ValorLiquido) AS Valor FROM #Rel2" & vbCrLf & _
                "GROUP BY Tipo, Grupo" & vbCrLf & _
                "ORDER BY Tipo" & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "SELECT rel.Tipo, isnull(rel.Grupo,'') AS Grupo, rel.Titulo, rel.Prorrogacao, rel.Empresa, rel.EndEmpresa, emp.Cidade, rel.Cliente, rel.EndCliente, rel.Nome, rel.Historico, rel.ValorLiquido" & vbCrLf & _
                "FROM #Rel1 rel" & vbCrLf & _
                "	INNER JOIN Clientes emp" & vbCrLf & _
                "			ON emp.Cliente_Id   = rel.Empresa" & vbCrLf & _
                "			AND emp.Endereco_Id = rel.EndEmpresa" & vbCrLf & _
                "ORDER BY rel.Tipo, rel.Empresa, rel.Titulo"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pagamentos")

        Return ds

    End Function

End Class