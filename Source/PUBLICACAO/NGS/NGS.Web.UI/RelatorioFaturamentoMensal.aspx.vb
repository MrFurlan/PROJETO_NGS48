Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class RelatorioFaturamentoMensal
    Inherits BasePage


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gerencial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioFaturamentoMensal", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaAno()
                CargaMes()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão acessar essa página.", "~/Gerencial.aspx")
                Exit Sub
            End If
        End If
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

    Private Sub CargaAno()
        Dim sql As String = "select distinct year(Movimento_Id)as Ano from Razao order by year(Movimento_Id) desc"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Ano")

        ddlAno.Items.Clear()

        For Each row As DataRow In ds.Tables(0).Rows
            ddlAno.Items.Add(New ListItem(row("Ano"), row("Ano")))
        Next
    End Sub

    Private Sub CargaMes()
        ddlMes.Items.Clear()

        For i = 1 To 12
            ddlMes.Items.Add(New ListItem(MonthName(i).ToUpper(), i))
        Next

        ddlMes.SelectedValue = DateTime.Now.Month
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Unidade: " & ddlUnidade.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlAno.SelectedValue) Then
            param &= "Ano: " & ddlAno.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlMes.SelectedValue) Then
            param &= "Mês: " & ddlMes.SelectedItem.Text & " - "
        End If

        Return param
    End Function

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            CargaEmpresas()
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
            If String.IsNullOrEmpty(ddlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a Empresa.")
                Exit Sub
            End If

            Dim ds As DataSet = getDataSet()
            For Each row As DataRow In ds.Tables(1).Rows
                If Not String.IsNullOrWhiteSpace(row("Cliente_Id")) Then
                    row("Cliente_Id") = Funcoes.FormatarCpfCnpj(row("Cliente_Id"))
                End If
            Next
            If Pdf Then
                Dim param As New Dictionary(Of String, Object)
                param.Add("Titulo", "Relatório De Faturamento Mensal.")
                param.Add("ConsultaParametros", getParam())
                param.Add("Consolidado", IIf(chkConsolidar.Checked, "(Consolidado)", ""))
                param.Add("Ano", ddlAno.SelectedValue)

                Funcoes.BindReport(Me.Page, ds, "Cr_Faturamento", eExportType.PDF, param)
            Else
                BindarExcel(ds, "Faturamento")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim whereConta As String = IIf(Left(ddlEmpresa.SelectedValue, 8) = "03189063", "", "or conta_id in('403010101', '301040106')")

        Dim sql As String = "SET LANGUAGE PORTUGUÊS; SELECT Mes, MesExtenso, Ano, SUM(Nacional) AS Nacional, SUM(Exterior) AS Exterior, SUM(Servicos) AS Servicos, SUM(MoedaNacional) AS NacionalMoeda, SUM(MoedaExterior)" & vbCrLf & _
                            "       AS ExteriorMoeda, SUM(MoedaServicos) AS ServicosMoeda, SUM(total) as Total, SUM(TotalExterior) as TotalMoeda" & vbCrLf & _
                            " FROM  (" & vbCrLf & _
                            " SELECT MONTH(Razao.Movimento_Id) as Mes, YEAR(Razao.Movimento_Id) AS Ano, UPPER(LEFT(DATENAME(month, Razao.Movimento_Id),1))+LOWER(SUBSTRING(DATENAME(MONTH, Razao.Movimento_Id),2,LEN(DATENAME(MONTH, Razao.Movimento_Id)))) AS MesExtenso, " & vbCrLf & _
                            "               ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7)= '3010101' " & whereConta & " THEN (CreditoOficial - DebitoOficial) ELSE 0 END), 0) AS Nacional, ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7)" & vbCrLf & _
                            "            = '3010102' THEN (CreditoOficial - DebitoOficial) ELSE 0 END), 0) AS Exterior, ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) IN ('3010103','3010104')" & vbCrLf & _
                            "            THEN (CreditoOficial - DebitoOficial) ELSE 0 END), 0) AS Servicos," & vbCrLf & _
                            "        ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) = '3010101' THEN (CreditoOficial - DebitoOficial) ELSE 0 END), 0) +                             " & vbCrLf & _
                            "            ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) = '3010102' THEN (CreditoOficial - DebitoOficial) ELSE 0 END), 0) +                         " & vbCrLf & _
                            "            ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) IN ('3010103','3010104') " & whereConta & " THEN (CreditoOficial - DebitoOficial) ELSE 0 END), 0) as Total,   " & vbCrLf & _
                            "            ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) = '3010101' " & whereConta & " THEN ((CreditoOficial / Indice)  - (DebitoOficial / indice)) ELSE 0 END), 0) AS MoedaNacional," & vbCrLf & _
                            "             ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) = '3010102' THEN ((CreditoOficial / indice) -  (DebitoOficial / Indice)) ELSE 0 END), 0) AS MoedaExterior," & vbCrLf & _
                            "             ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) IN ('3010103','3010104') THEN ((CreditoOficial / indice) - (DebitoOficial / Indice)) ELSE 0 END), 0) AS MoedaServicos," & vbCrLf & _
                            "        ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) = '3010101' THEN ((CreditoOficial / Indice)  - (DebitoOficial / indice)) ELSE 0 END), 0) +                                    " & vbCrLf & _
                            "             ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) = '3010102' THEN ((CreditoOficial / indice) -  (DebitoOficial / Indice)) ELSE 0 END), 0) +                               " & vbCrLf & _
                            "             ISNULL(SUM(CASE WHEN LEFT(Conta_ID, 7) IN ('3010103','3010104') " & whereConta & " THEN ((CreditoOficial / indice) - (DebitoOficial / Indice)) ELSE 0 END), 0) as TotalExterior   " & vbCrLf & _
                            "       FROM  Razao INNER JOIN" & vbCrLf & _
                            "            Cotacoes ON Razao.Movimento_Id = Cotacoes.Data_Id AND Cotacoes.Indexador_Id = 3" & vbCrLf

        If chkConsolidar.Checked Then
            sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
        Else
            sql = sql & " WHERE Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf
        End If

        sql &= "AND Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
               " And Lote_Id not in (7500) And  (Left(Conta_ID, 7) in ('3010101', '3010102', '3010103', '3010104') or conta_id = '301040106' " & whereConta & ") And Year(Movimento_Id) = " & ddlAno.SelectedValue & " And month(Movimento_ID) <= " & Left(ddlMes.SelectedValue, 2) & vbCrLf & _
               " Group By Razao.Movimento_Id ) as Consulta Group By Ano, Mes, MesExtenso order by Mes" & vbCrLf & _
               ""
        sql &= "select  nome, Cidade, Estado, Cliente_Id  from Clientes where Cliente_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' and Endereco_Id = " & ddlEmpresa.SelectedValue.Split("-")(1)

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Faturamento")
        ds.Tables(1).TableName = "Clientes"

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing Then
            Dim Acum As Decimal = 0
            Dim acumEx As Decimal = 0

            ds.Tables(0).Columns.Add("Acumulado", Type.GetType("System.Decimal"))
            ds.Tables(0).Columns.Add("AcumuladoMoeda", Type.GetType("System.Decimal"))

            For Each row As DataRow In ds.Tables(0).Rows
                row("Acumulado") = row("Total") + Acum
                row("AcumuladoMoeda") = row("TotalMoeda") + acumEx
                Acum = row("Acumulado")
                acumEx = row("AcumuladoMoeda")
            Next
        End If

        Return ds
    End Function

    Private Sub BindarExcel(ByVal ds As DataSet, ByVal TituloAba As String)
        Try
            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                Dim rowIndexDolar = ds.Tables(0).Rows.Count + 9
                Dim rowIndexDolarSoma = rowIndexDolar + 1
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1
                Try
                    Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    If File.Exists(fileName) Then File.Delete(fileName)

                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)

                            'criando aba da planilha.
                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                            'Inserindo Cabeçalho Empresa e Cidade e Título
                            worksheet.Cells(1, 1, 5, 6).Style.Font.Bold = True
                            worksheet.Cells(1, 1, 1, 6).Merge = True
                            worksheet.Cells(2, 1, 2, 6).Merge = True
                            worksheet.Cells(4, 2, 4, 3).Merge = True
                            worksheet.Cells(5, 5, 5, 6).Merge = True

                            worksheet.Cells(rowIndex, columnIndex).Value = " " & ddlEmpresa.SelectedItem.Text.Split("-")(0).Replace("..", "")
                            rowIndex += 1

                            worksheet.Cells(rowIndex, columnIndex).Value = IIf(Not chkConsolidar.Checked, ddlEmpresa.SelectedItem.Text.Split("-")(1), "")
                            rowIndex += 2

                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Faturamento " & ddlAno.SelectedValue & " " & IIf(chkConsolidar.Checked, "(Consolidado)", "")
                            rowIndex += 1
                            columnIndex = 5
                            worksheet.Cells(rowIndex, columnIndex).Value = "Emissão: " & DateTime.Now

                            columnIndex = 1

                            'Inserindo as colunas.
                            worksheet.Cells(rowIndex, columnIndex).Value = "Em Reais"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Em Dolar"
                            worksheet.Cells(rowIndexDolar, columnIndex).Style.Font.Bold = True
                            rowIndex += 1
                            rowIndexDolar += 1

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, columnIndex, rowIndex, ds.Tables(0).Columns.Count - 7)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            Using range = worksheet.Cells(rowIndexDolar, columnIndex, rowIndexDolar, ds.Tables(0).Columns.Count - 7)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            worksheet.Cells(rowIndex, columnIndex).Value = "Mês"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Mês"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Nacional"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Nacional"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Exterior"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Exterior"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Serviços"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Serviços"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Total"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Total"
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Acumulado"
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "Acumulado"

                            rowIndex += 1
                            rowIndexDolar += 1

                            Dim primeiraLinha As Boolean = True

                            For Each row As DataRow In ds.Tables(0).Rows

                                'Formatacoes de celulas
                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 7)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If

                                If rowIndexDolar Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndexDolar, 1, rowIndexDolar, ds.Tables(0).Columns.Count - 7)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If
                                columnIndex = 1
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    If col.ColumnName = "MesExtenso" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row("MesExtenso")
                                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                        worksheet.Cells(rowIndexDolar, columnIndex).Value = row("MesExtenso")
                                        worksheet.Cells(rowIndexDolar, columnIndex).Style.Font.Bold = True
                                        columnIndex += 1
                                    ElseIf col.ColumnName = "Nacional" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row("Nacional")
                                        worksheet.Cells(rowIndexDolar, columnIndex).Value = row("NacionalMoeda")
                                        columnIndex += 1
                                    ElseIf col.ColumnName = "Exterior" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row("Exterior")
                                        worksheet.Cells(rowIndexDolar, columnIndex).Value = row("ExteriorMoeda")
                                        columnIndex += 1
                                    ElseIf col.ColumnName = "Servicos" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row("Servicos")
                                        worksheet.Cells(rowIndexDolar, columnIndex).Value = row("ServicosMoeda")
                                        columnIndex += 1
                                    End If

                                Next

                                'For columnIndex = 1 To 4
                                '    If columnIndex = 1 Then
                                '        worksheet.Cells(rowIndex, columnIndex).Value = MonthName(row(columnIndex - 1))
                                '        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                '        worksheet.Cells(rowIndexDolar, columnIndex).Value = MonthName(row(columnIndex - 1))
                                '        worksheet.Cells(rowIndexDolar, columnIndex).Style.Font.Bold = True
                                '    Else
                                '        worksheet.Cells(rowIndex, columnIndex).Value = row(columnIndex)
                                '        worksheet.Cells(rowIndexDolar, columnIndex).Value = row(columnIndex + 3)
                                '    End If
                                'Next

                                worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(rowIndex, 2).Address & ":" & worksheet.Cells(rowIndex, 4).Address & ")"
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells(rowIndexDolar, columnIndex).Formula = "SUM(" & worksheet.Cells(rowIndexDolar, 2).Address & ":" & worksheet.Cells(rowIndexDolar, 4).Address & ")"
                                worksheet.Cells(rowIndexDolar, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                If primeiraLinha Then
                                    worksheet.Cells(rowIndex, columnIndex + 1).Formula = "SUM(" & worksheet.Cells(rowIndex, 5).Address & ")"
                                    worksheet.Cells(rowIndexDolar, columnIndex + 1).Formula = "SUM(" & worksheet.Cells(rowIndexDolar, 5).Address & ")"
                                    primeiraLinha = False
                                Else
                                    worksheet.Cells(rowIndex, columnIndex + 1).Formula = "SUM(" & worksheet.Cells(rowIndex, 5).Address & " + " & worksheet.Cells(rowIndex - 1, 6).Address & ")"
                                    worksheet.Cells(rowIndexDolar, columnIndex + 1).Formula = "SUM(" & worksheet.Cells(rowIndexDolar, 5).Address & " + " & worksheet.Cells(rowIndexDolar - 1, 6).Address & ")"
                                End If

                                rowIndex += 1
                                rowIndexDolar += 1
                            Next

                            columnIndex = 1

                            'aplicando formatação nas células do RODAPÉ
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 7)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            Using range = worksheet.Cells(rowIndexDolar, 1, rowIndexDolar, ds.Tables(0).Columns.Count - 7)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL..."
                            worksheet.Cells(rowIndexDolar, columnIndex).Value = "TOTAL..."

                            For columnIndex = 2 To 6
                                If columnIndex = 6 Then
                                    worksheet.Cells(rowIndex, columnIndex).Formula = worksheet.Cells(rowIndex - 1, columnIndex).Formula
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                    worksheet.Cells(rowIndexDolar, columnIndex).Formula = worksheet.Cells(rowIndexDolar - 1, columnIndex).Formula
                                    worksheet.Cells(rowIndexDolar, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(1, columnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, columnIndex).Address & ")"
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                    worksheet.Cells(rowIndexDolar, columnIndex).Formula = "SUM(" & worksheet.Cells(rowIndexDolarSoma, columnIndex).Address & ":" & worksheet.Cells(rowIndexDolar - 1, columnIndex).Address & ")"
                                    worksheet.Cells(rowIndexDolar, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                End If

                            Next

                            worksheet.Cells(6, 1, rowIndexDolar, columnIndex + 2).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            worksheet.Cells.AutoFitColumns(0)

                            worksheet.View.FreezePanes(3, 1)

                            worksheet.Column(1).Width = 14
                            worksheet.Column(2).Width = 14
                            worksheet.Column(3).Width = 14
                            worksheet.Column(4).Width = 14
                            worksheet.Column(5).Width = 14
                            worksheet.Column(6).Width = 14

                            package.Save()
                        End Using
                    End Using

                    'download do arquivo pelo browser
                    Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                Catch ex As Exception
                    Throw New Exception(ex.Message & " - linha: " & rowIndex & " - coluna: " & columnIndex)
                End Try

            Else
                MsgBox(Me.Page, "Nenhum resultado encontrado.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioFaturamentoMensal")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class