Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class RelatorioDeAnaliseDeCredito
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                'If Funcoes.VerificaPermissao("RelatorioDeAnaliseDeCredito", "ACESSAR") Then
                CargaUnidade()
                    VerificaUnidade()
                    CargaAno()
                    CargaRepresentante()
                    LiberaEmpresa()
                'Else
                '    If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Financeiro)
                '    Exit Sub
                'End If
            End If
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

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub CargaRepresentante()
        Dim ds As DataSet
        Dim strRepre As String = ""
        Dim codRep As String = ""

        Dim Sql As String = "  SELECT distinct    Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf & _
                            "   FROM   Pedidos INNER JOIN" & vbCrLf & _
                            "          Clientes ON Pedidos.Cliente = Clientes.Cliente_Id AND Pedidos.EndCliente = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
                            "          SubOperacoes ON Pedidos.Operacao = SubOperacoes.Operacao_Id AND Pedidos.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                            "   Where SubOperacoes.EntradaSaida = 'S'" & vbCrLf & _
                            "   Order by Nome" & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "Cliente")

        ddlCliente.Items.Clear()

        For Each drCliente As DataRow In ds.Tables(0).Rows
            strRepre = Funcoes.AlinharEsquerda(drCliente("Nome").ToString(), 50, ".") & " - " & drCliente("Cliente_Id") & "-" & drCliente("Endereco_Id")
            codRep = drCliente("Cliente_Id").ToString() & "-" & drCliente("Endereco_Id")
            ddlCliente.Items.Add(New ListItem(strRepre, codRep))
        Next
        ddlCliente.Items.Insert(0, "")
    End Sub

    Private Sub CargaAno()
        Dim sql As String = "select distinct year(Movimento_Id)as Ano from Razao order by year(Movimento_Id) desc"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Ano")

        ddlAno.Items.Clear()

        For Each row As DataRow In ds.Tables(0).Rows
            ddlAno.Items.Add(New ListItem(row("Ano"), row("Ano")))
        Next
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = ""
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "")

        Return ds
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim rowIndex As Integer = 1
            Dim ColumnIndex As Integer = 1
            Dim rowInicial As Integer = 1

            Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then File.Delete(fileName)

            'emitir excel.xsls do office / relatório padrão em lista
            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando aba da planilha.
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("AnaliseDeCredito")

                    Dim objEmpresa As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))

                    'Inserindo o cabeçalho.
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objEmpresa.Nome
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objEmpresa.Cidade & " - " & objEmpresa.Estado.Codigo
                    rowIndex += 2
                    worksheet.Cells(rowIndex, ColumnIndex).Value = "Análise de Crédito"
                    rowIndex += 2
                    rowInicial = rowIndex

                    'Formatacao do cabeçalho
                    worksheet.Cells(1, 1, 22, 5).Style.Font.Bold = True

                    Dim cab As String() = {"Cliente", "Nome", "Endereco", "Número", "Complemento", "Cidade", "Estado", _
                                           "Inscrição", "Telefone", "E-Mail", "", "Aquisições#Limite de Crédito", "Liquidações#Limite Utilizado", _
                                           "Pendências#Saldo", "", "Média de Atraso#Menor Valor", "Maior Atraso#Maior Valor"}

                    For Each item As String In cab
                        If Not item.Contains("#") Then
                            worksheet.Cells(rowIndex, ColumnIndex).Value = item
                        Else
                            worksheet.Cells(rowIndex, ColumnIndex).Value = item.Split("#")(0)
                            worksheet.Cells(rowIndex, ColumnIndex + 3).Value = item.Split("#")(1)
                        End If
                        rowIndex += 1
                    Next

                    ColumnIndex += 1
                    rowIndex = rowInicial

                    Dim objCliente = New Cliente(ddlCliente.SelectedValue.Split("-")(0), ddlCliente.SelectedValue.Split("-")(1))

                    'informações do cliente
                    worksheet.Cells(rowIndex, ColumnIndex).Value = Funcoes.FormatarCpfCnpj(objCliente.Codigo)
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Nome
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Endereco
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Numero
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Complemento
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Cidade
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Estado
                    rowIndex += 1

                    If Not String.IsNullOrWhiteSpace(objCliente.InscricaoEstadual) Then
                        If IsNumeric(objCliente.InscricaoEstadual) Then
                            worksheet.Cells(rowIndex, ColumnIndex).Value = Convert.ToInt64(objCliente.InscricaoEstadual)
                            worksheet.Cells(rowIndex, ColumnIndex + 3).Style.Numberformat.Format = "0"
                        Else
                            worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.InscricaoEstadual
                        End If
                    End If
                    rowIndex += 1

                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Telefone
                    rowIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = objCliente.Email
                    rowIndex += 2

                    Dim Provisao As String = "1, 2"

                    If FinanceiroNovo() Then Provisao = "1, 3"

                    Dim ds As DataSet = getDataSet(Provisao)
                    ds.Merge(getDataSet("1"))
                    ds.Merge(getDataSetMP_MA("MediaPgto"))
                    ds.Merge(getDataSetMP_MA("MaiorAtraso"))
                    ds.Merge(getDataSetMValor("MaiorValor"))
                    ds.Merge(getDataSetMValor("MenorValor"))
                    ds.Merge(getDataSetTiTulosPendentes())

                    'Aquisicoes
                    For Each row As DataRow In ds.Tables("Aquisicoes").Rows
                        worksheet.Cells(rowIndex, ColumnIndex).Value = row("Aquisicoes")
                    Next

                    rowIndex += 1

                    'Liquidacoes
                    For Each row As DataRow In ds.Tables("Liquidacoes").Rows
                        worksheet.Cells(rowIndex, ColumnIndex).Value = row("Liquidacoes")
                    Next
                    rowIndex += 1

                    worksheet.Cells(rowIndex, ColumnIndex).Formula = "SUM(" & worksheet.Cells(rowIndex - 2, ColumnIndex).Address & "-" & worksheet.Cells(rowIndex - 1, ColumnIndex).Address & ")"
                    worksheet.Cells(rowIndex - 2, ColumnIndex, rowIndex, ColumnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    rowIndex += 2

                    'MediaPgto
                    For Each row As DataRow In ds.Tables("MediaPgto").Rows
                        If row("Total") <> 0 And row("Valor") <> 0 Then
                            worksheet.Cells(rowIndex, ColumnIndex).Value = Format(row("Total") / (row("Valor")), "#,##0.00") & " Dias"
                        Else
                            worksheet.Cells(rowIndex, ColumnIndex).Value = "0 Dias"
                        End If
                    Next

                    'Menor Valor
                    For Each row As DataRow In ds.Tables("MenorValor").Rows
                        worksheet.Cells(rowIndex, ColumnIndex + 3).Value = row("Valor")
                        worksheet.Cells(rowIndex, ColumnIndex + 3).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    Next
                    rowIndex += 1

                    'MaiorAtraso
                    For Each row As DataRow In ds.Tables("MaiorAtraso").Rows
                        If row("Total") <> 0 And row("Valor") <> 0 Then
                            worksheet.Cells(rowIndex, ColumnIndex).Value = Format(row("Total") / (row("Valor")), "#,##0.00") & " Dias"
                        Else
                            worksheet.Cells(rowIndex, ColumnIndex).Value = "0 Dias"
                        End If
                    Next

                    'Maior Valor
                    For Each row As DataRow In ds.Tables("MaiorValor").Rows
                        worksheet.Cells(rowIndex, ColumnIndex + 3).Value = row("Valor")
                        worksheet.Cells(rowIndex, ColumnIndex + 3).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    Next
                    rowIndex += 2
                    ColumnIndex = 1

                    'formata cabeçalho titulos pendentes
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex + 1, ds.Tables("TitulosPendentes").Columns.Count + 3)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using

                    'Titulos Pendentes
                    worksheet.Cells(rowIndex, ColumnIndex).Value = "Titulos Pendentes"
                    rowIndex += 1

                    Dim i As Integer

                    For Each col As DataColumn In ds.Tables("TitulosPendentes").Columns
                        If i = 0 Then
                            worksheet.Cells(rowIndex, ColumnIndex).Value = col.ColumnName
                            worksheet.Cells(rowIndex, 1, rowIndex, 4).Merge = True
                            ColumnIndex = 5
                            i = 1
                        Else
                            worksheet.Cells(rowIndex, ColumnIndex).Value = col.ColumnName
                            ColumnIndex += 1
                        End If
                    Next
                    rowIndex += 1

                    Dim rowPendentes As Integer = rowIndex
                    worksheet.Cells(rowIndex, 1, rowIndex, 4).Merge = True

                    For Each row As DataRow In ds.Tables("TitulosPendentes").Rows
                        ColumnIndex = 1
                        i = 0
                        For Each col As DataColumn In ds.Tables("TitulosPendentes").Columns
                            If i = 0 Then
                                worksheet.Cells(rowIndex, 1, rowIndex, 4).Merge = True
                                worksheet.Cells(rowIndex, ColumnIndex).Value = row(col.ColumnName)
                                ColumnIndex = 5
                                i = 1
                            Else
                                If IsDate(row(col.ColumnName)) Then
                                    worksheet.Cells(rowIndex, ColumnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                ElseIf IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString().Contains(",") Then
                                    worksheet.Cells(rowIndex, ColumnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                End If

                                worksheet.Cells(rowIndex, ColumnIndex).Value = row(col.ColumnName)
                                ColumnIndex += 1
                            End If
                        Next
                        rowIndex += 1
                    Next
                    ColumnIndex = 1
                    worksheet.Cells(rowIndex, ColumnIndex).Value = "Total"
                    worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("TitulosPendentes").Columns.Count + 3).Style.Font.Bold = True
                    ColumnIndex += 7

                    worksheet.Cells(rowIndex, ColumnIndex).Formula = "SUM(" & worksheet.Cells(rowPendentes, ColumnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, ColumnIndex).Address & ")"
                    worksheet.Cells(rowIndex, ColumnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    ColumnIndex += 1
                    worksheet.Cells(rowIndex, ColumnIndex).Formula = "SUM(" & worksheet.Cells(rowPendentes, ColumnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, ColumnIndex).Address & ")"
                    worksheet.Cells(rowIndex, ColumnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'formata cabeçalho titulos pendentes
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables("TitulosPendentes").Columns.Count + 3)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using

                    ColumnIndex += 1

                    'criando auto filtro na planilha
                    'worksheet.Cells(rowPendentes - 1, 1, rowPendentes - 1, ColumnIndex).AutoFilter = True

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    worksheet.Column(1).Width = 20

                    'congelando primeira linham
                    worksheet.View.FreezePanes(rowInicial - 1, 1)

                    'salvando planilha do excel
                    package.Save()

                End Using
            End Using

            'download do arquivo pelo browser
            Funcoes.AbrirExcel(Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet(ByVal provisao As String)
        Try
            Dim sql As String = ""

            If FinanceiroNovo Then
                sql = " Select SUM(ValorOficial) " & IIf(provisao = "1", "Liquidacoes", "Aquisicoes") & vbCrLf & _
                      "    from Titulos t                          " & vbCrLf & _
                      "   Inner Join TitulosxContaContabil tc      " & vbCrLf & _
                      "      on tc.Titulo_Id = t.Titulo_Id         " & vbCrLf & _
                      " 	and tc.Conta_Id  = t.ContaContabilCliFor" & vbCrLf & _
                      " Where   t.RecPag = 'R'                   " & vbCrLf & _
                      " 	And t.CliFor   = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                      " 	And t.Situacao = 1                     " & vbCrLf & _
                      " 	And t.Provisao in (" & provisao & ")                " & vbCrLf & _
                      "  And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf
            Else
                sql = " Select Sum(ValorDoDocumento) as " & IIf(provisao = "1", "Liquidacoes", "Aquisicoes") & " from ContasAReceber" & vbCrLf & _
                      "     where Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'  And Situacao = 1 And Provisao in (" & provisao & ")" & vbCrLf & _
                      "         And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf
            End If

            Return Banco.ConsultaDataSet(sql, IIf(provisao = "1", "Liquidacoes", "Aquisicoes"))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetMP_MA(ByVal table As String)
        Dim sql As String = ""
        Try
            If FinanceiroNovo Then
                sql &= "   SELECT     SUM(Valor * dias) AS Total, SUM(Dias) AS Dias, Sum(Valor) as Valor from (   " & vbCrLf & _
                       "   Select tc.ValorOficial as Valor, Sum(DATEDIFF(DAY, t.Vencimento, t.DataBaixa)) As Dias " & vbCrLf & _
                       "    from Titulos t                                                                      " & vbCrLf & _
                       "        Inner Join TitulosxContaContabil tc                                                  " & vbCrLf & _
                       "            on tc.Titulo_Id = t.Titulo_Id                                                     " & vbCrLf & _
                       "   	        and tc.Conta_Id = t.ContaContabilCliFor                                            " & vbCrLf & _
                       "    Where t.RecPag = 'R'                                                                   " & vbCrLf & _
                       "   	    And t.Situacao = 1                                                                 " & vbCrLf & _
                       "   	    and t.Provisao = 1                                                                 " & vbCrLf & _
                       "   	    And t.CliFor = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                       "   	    And year(t.Movimento) >= " & ddlAno.SelectedValue & vbCrLf & _
                       "   Group by tc.ValorOficial                                                               " & vbCrLf & _
                       "                                                                                          " & vbCrLf & _
                       "   Union All                                                                              " & vbCrLf & _
                       "                                                                                          " & vbCrLf & _
                       "   Select tc.ValorOficial as Valor, Sum(DATEDIFF(DAY, t.Vencimento, '" & Now.ToSqlDate() & "')) As Dias" & vbCrLf & _
                       "    from Titulos t                                                                      " & vbCrLf & _
                       "        Inner Join TitulosxContaContabil tc                                                  " & vbCrLf & _
                       "            on tc.Titulo_Id = t.Titulo_Id                                                     " & vbCrLf & _
                       "   	        and tc.Conta_Id = t.ContaContabilCliFor                                            " & vbCrLf & _
                       "    Where t.RecPag = 'R'                                                                   " & vbCrLf & _
                       "   	    And t.Situacao = 1                                                                 " & vbCrLf & _
                       "   	    and t.Provisao = 3                                                                 " & vbCrLf & _
                       "   	    And t.CliFor = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                       "   	    And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf & _
                       "   	    And t.Vencimento <= '" & Now.ToSqlDate() & "'" & vbCrLf & _
                       "   Group by tc.ValorOficial) as Consulta Where Dias " & IIf(table = "MediaPgto", ">", "<>") & " 0" & vbCrLf
            Else
                sql = "   SELECT     SUM(Valor * dias) AS Total, SUM(Dias) AS Dias, Sum(Valor) as Valor from ( " & vbCrLf & _
                      "   Select ValorDoDocumento As Valor, Sum(DATEDIFF(DAY," & vbCrLf & _
                      "        Vencimento,Baixa)) As Dias From ContasAReceber" & vbCrLf & _
                      "        Where Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "' And Situacao = 1 And Provisao = 1" & vbCrLf & _
                      " And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf & _
                      " Group By ValorDoDocumento " & vbCrLf & _
                      " " & vbCrLf & _
                      " Union" & vbCrLf & _
                      " Select ValorDoDocumento As Valor, Sum(DATEDIFF(DAY," & vbCrLf & _
                      "     Vencimento, '" & Now.ToSqlDate() & "')) As Dias From ContasAReceber" & vbCrLf & _
                      "     Where Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "' And Situacao = 1 And Provisao = 2" & vbCrLf & _
                      " And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf & _
                      " And (Vencimento <= '" & Now.ToSqlDate() & "')" & vbCrLf & _
                      " Group By ValorDoDocumento " & vbCrLf & _
                      "    " & vbCrLf & _
                      " ) as Consulta Where Dias " & IIf(table = "MediaPgto", ">", "<>") & " 0" & vbCrLf
            End If
            Return Banco.ConsultaDataSet(sql, table)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetMValor(ByVal tipo As String)
        Dim sql As String = ""
        Try
            If FinanceiroNovo Then
                sql &= "    Select " & IIf(tipo = "MenorValor", "Min", "Max") & "(tc.ValorOficial) as Valor" & vbCrLf & _
                    "      from Titulos t                           " & vbCrLf & _
                    "     Inner Join TitulosxContaContabil tc       " & vbCrLf & _
                    "        on tc.Titulo_Id = t.Titulo_Id          " & vbCrLf & _
                    "   	and tc.Conta_Id = t.ContaContabilCliFor " & vbCrLf & _
                    "   Where t.RecPag = 'R'                        " & vbCrLf & _
                    "   	And t.Situacao = 1                      " & vbCrLf & _
                    "   	and t.Provisao in (1, 3)                " & vbCrLf & _
                    "   	And t.CliFor = '" & ddlCliente.SelectedValue.Split("-")(0) & "'         " & vbCrLf & _
                    "   	And year(t.Movimento) >= " & ddlAno.SelectedValue & vbCrLf
            Else
                sql &= " Select " & IIf(tipo = "MenorValor", "Min", "Max") & "(ValorDoDocumento) as Valor from ContasAReceber" & vbCrLf & _
                       " where Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'  And Situacao = 1 And Provisao in (1, 2)" & vbCrLf & _
                       " And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf
            End If

            Return Banco.ConsultaDataSet(sql, tipo)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetTiTulosPendentes()
        Dim sql As String = ""
        Try
            If FinanceiroNovo Then
                sql &= "   Select t.Historico, t.Pedido, t.Titulo_Id as Titulo, t.Reprogramacao as Vencimento,       " & vbCrLf & _
                    "     case when t.Vencimento < GETDATE()                                                         " & vbCrLf & _
                    "      	Then tc.ValorOficial                                                                     " & vbCrLf & _
                    "      	else 0                                                                                   " & vbCrLf & _
                    "      end as Vencidos,                                                                          " & vbCrLf & _
                    "      case when t.Vencimento < GETDATE()                                                        " & vbCrLf & _
                    "      	Then 0                                                                                   " & vbCrLf & _
                    "      	else tc.ValorOficial                                                                     " & vbCrLf & _
                    "      end as 'A Vencer',                                                                        " & vbCrLf & _
                    "      case when Sum(DATEDIFF(DAY, t.Vencimento, GETDATE())) > 0                                 " & vbCrLf & _
                    "      	then  Sum(DATEDIFF(DAY, t.Vencimento, GETDATE()))                                        " & vbCrLf & _
                    "      	else 0                                                                                   " & vbCrLf & _
                    "      end As Dias                                                                               " & vbCrLf & _
                    "      from Titulos t                                                                            " & vbCrLf & _
                    "     Inner Join TitulosxContaContabil tc                                                        " & vbCrLf & _
                    "        on tc.Titulo_Id = t.Titulo_Id                                                           " & vbCrLf & _
                    "   	and tc.Conta_Id = t.ContaContabilCliFor                                                  " & vbCrLf & _
                    "   Where t.RecPag = 'R'                                                                         " & vbCrLf & _
                    "   	And t.Situacao = 1                                                                       " & vbCrLf & _
                    "   	and t.Provisao = 3                                                                       " & vbCrLf & _
                    "   	And t.CliFor = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                    "   	And year(t.Movimento) >= " & ddlAno.SelectedValue & vbCrLf & _
                    "   Group By t.Historico, t.Pedido, t.Titulo_Id, t.Reprogramacao,   tc.ValorOficial, t.Vencimento" & vbCrLf & _
                    "    Order By t.Reprogramacao                                                                    " & vbCrLf
            Else
                sql &= "   Select	Historico, Pedido, Registro_Id as Titulo, Prorrogacao as Vencimento,  " & vbCrLf & _
                       "   case when Vencimento < GETDATE()                                              " & vbCrLf & _
                       "   	Then ValorDoDocumento                                                     " & vbCrLf & _
                       "   	else 0                                                                    " & vbCrLf & _
                       "   end as Vencidos,                                                              " & vbCrLf & _
                       "   case when Vencimento < GETDATE()                                              " & vbCrLf & _
                       "   	Then 0                                                                    " & vbCrLf & _
                       "   	else ValorDoDocumento                                                     " & vbCrLf & _
                       "   end as 'A Vencer',                                                            " & vbCrLf & _
                       "                                                                                 " & vbCrLf & _
                       "   case when Sum(DATEDIFF(DAY, Vencimento, GETDATE())) > 0                     " & vbCrLf & _
                       "   	then  Sum(DATEDIFF(DAY, Vencimento, GETDATE()))                         " & vbCrLf & _
                       "   	else 0                                                                    " & vbCrLf & _
                       "   end As Dias                                                                   " & vbCrLf & _
                       " From ContasAReceber" & vbCrLf & _
                       " where Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                       " And Situacao = 1" & vbCrLf & _
                       " And year(Movimento) >= " & ddlAno.SelectedValue & vbCrLf & _
                       " And Provisao = 2" & vbCrLf & _
                       " Group By Historico, Pedido, Registro_Id, Prorrogacao,   ValorDoDocumento, Vencimento" & vbCrLf & _
                       " Order By Prorrogacao" & vbCrLf
            End If

            Return Banco.ConsultaDataSet(sql, "TitulosPendentes")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeAnaliseDeCredito")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class