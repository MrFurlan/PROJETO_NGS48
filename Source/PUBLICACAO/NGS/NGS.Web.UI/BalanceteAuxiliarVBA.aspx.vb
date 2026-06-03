Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class BalanceteAuxiliarVBA
    Inherits BasePage

    Dim SdoAnterior As Double
    Dim Debitos As Double
    Dim Creditos As Double
    Dim GConta As String
    Dim GCliente As String
    Dim GEndCliente As String
    Dim GSaldoTitulos As Double

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("BalanceteAuxiliarVBA", "ACESSAR") Then
                    limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampo() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) AndAlso String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período de consulta.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe a data inicial.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe a data final.")
            Return False
        ElseIf Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Data inicial não é uma data válida.")
            Return False
        ElseIf Not IsDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data final não é uma data válida.")
            Return False
        ElseIf Year(txtDataInicial.Text) <> Year(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe as datas no mesmo ano.")
            Return False
        End If
        Return True
    End Function

    Private Function getDataSetPlanilha() As DataSet
        Dim Sql As String = "SELECT Empresa.Reduzido + ' - ' + LEFT(Empresa.Nome, 10) + ' - ' + Empresa.Cidade AS Empresa," & vbCrLf & _
                            "       A.Conta_Id as Conta," & vbCrLf & _
                            "       PlanoDeContas.Titulo AS TituloDaConta," & vbCrLf & _
                            "       A.Cliente_Id as Cliente," & vbCrLf & _
                            "       ISNULL(Clientes.Nome, N'') AS NomeDoCliente," & vbCrLf & _
                            "       A.Movimento_Id as Movimento," & vbCrLf & _
                            "       A.Lote_Id as Lote," & vbCrLf & _
                            "       A.Sequencia_Id as Sequencia," & vbCrLf & _
                            "       A.Titulo," & vbCrLf & _
                            "       A.Pedido," & vbCrLf & _
                            "       A.Produto," & vbCrLf & _
                            "       Produtos.Nome as NomeDoProduto," & vbCrLf & _
                            "       Case When Origem = 'Financeiro' then A.DebitoOficial  else 0 end as DebitoAdiantamento," & vbCrLf & _
                            "       Case When Origem = 'Financeiro' then A.CreditoOficial else 0 end as CreditoAdiantamento," & vbCrLf & _
                            "       Case When Origem = 'Razao'      then A.DebitoOficial  else 0 end as DebitoAmortizacao," & vbCrLf & _
                            "       Case When Origem = 'Razao'      then A.CreditoOficial else 0 end as CreditoAmortizacao," & vbCrLf & _
                            "       Case When Origem = 'Notas'      then A.DebitoOficial  else 0 end as DebitoNotas," & vbCrLf & _
                            "       Case When Origem = 'Notas'      then A.CreditoOficial else 0 end as CreditoNotas," & vbCrLf & _
                            "       A.Historico," & vbCrLf & _
                            "       A.DataDaNota," & vbCrLf & _
                            "       A.Origem" & vbCrLf & _
                            "  FROM Gerencial.dbo.vw_Adiantamento A" & vbCrLf & _
                            " INNER JOIN Clientes AS Empresa" & vbCrLf & _
                            "    ON A.Empresa_Id    = Empresa.Cliente_Id" & vbCrLf & _
                            "   AND A.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf & _
                            "  LEFT JOIN Produtos" & vbCrLf & _
                            "    ON A.Produto = Produtos.Produto_Id " & vbCrLf & _
                            "  LEFT JOIN Clientes" & vbCrLf & _
                            "    ON A.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf & _
                            "   AND A.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
                            "  LEFT JOIN PlanoDeContas" & vbCrLf & _
                            "    ON A.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf & _
                            " Order By Empresa.Reduzido, A.Cliente_Id, A.Conta_Id, Movimento_Id" & vbCrLf
        Return Banco.ConsultaDataSet(Sql, "MercadoriasEmTransito")
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidaCampo() Then
                If chkPlanilha.Checked Then
                    Dim dsPlanilha As DataSet = getDataSetPlanilha()

                    If dsPlanilha IsNot Nothing AndAlso dsPlanilha.Tables IsNot Nothing AndAlso dsPlanilha.Tables(0).Rows.Count > 0 Then

                        Dim colunas As New Dictionary(Of String, eTipoCampo)
                        colunas.Add("Movimento", eTipoCampo.Data)
                        colunas.Add("DebitoAdiantamento", eTipoCampo.ValorSemTotalizador)
                        colunas.Add("CreditoAdiantamento", eTipoCampo.ValorSemTotalizador)
                        colunas.Add("DebitoAmortizacao", eTipoCampo.ValorSemTotalizador)
                        colunas.Add("CreditoAmortizacao", eTipoCampo.ValorSemTotalizador)
                        colunas.Add("DebitoNotas", eTipoCampo.ValorSemTotalizador)
                        colunas.Add("CreditoNotas", eTipoCampo.ValorSemTotalizador)

                        Funcoes.BindExcelOffice(Me.Page, dsPlanilha, "Mercadorias Em Trânsito", colunas)

                    Else
                        MsgBox(Me.Page, "Nenhum resultado encontrado.")
                    End If
                Else
                    Dim Anterior As String = ""
                    Dim ds As DataSet = getDataSetCompleto()

                    If ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Nenhum resultado encontrado!")
                    Else
                        'Emitir Excel.xsls do office/ Relatorio padrao em lista.
                        Dim rowIndex As Integer = 1
                        Dim columnIndex As Integer = 1

                        Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                        If File.Exists(fileName) Then File.Delete(fileName)

                        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                            Using package As New ExcelPackage(arquivo)

                                'criando aba da planilha.
                                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Balancete")

                                'Inserindo o Cabeçalho.
                                Dim objEmpresa As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                                worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Nome.ToUpper
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                rowIndex += 1
                                columnIndex += 1
                                worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Cidade.ToUpper & "/" & objEmpresa.Estado.Descricao.ToUpper
                                rowIndex += 1
                                worksheet.Cells(rowIndex, columnIndex).Value = "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo) & "-" & objEmpresa.Endereco
                                rowIndex += 2

                                'Inserindo o Título
                                If rdAuxiliar.Checked Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "Balancete Auxiliar"
                                ElseIf rdGerencial.Checked Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "Balancete Gerencial Financeiro"
                                ElseIf rdAdiantamentos.Checked Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "Balancete de Adiantamentos"
                                End If
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                rowIndex += 2

                                'Inserindo informacoes adicionais
                                columnIndex = 1
                                worksheet.Cells(rowIndex, columnIndex).Value = "Período: " & txtDataInicial.Text & " À " & txtDataFinal.Text
                                columnIndex = ds.Tables(0).Columns.Count - 2
                                worksheet.Cells(rowIndex, columnIndex).Value = "Emissão: " & DateTime.Now
                                rowIndex += 1

                                worksheet.Cells("A8").Value = "Conta"
                                worksheet.Cells("B8").Value = "Titulo"
                                worksheet.Cells("C8").Value = "Saldo Anterior"
                                worksheet.Cells("D8").Value = "Débitos"
                                worksheet.Cells("E8").Value = "Créditos"
                                worksheet.Cells("F8").Value = "Saldo Final"
                                worksheet.Cells("H8").Value = "Saldo Titulos"
                                worksheet.Cells("I8").Value = "Diferença"

                                rowIndex += 1

                                For Each row As DataRow In ds.Tables(0).Rows
                                    If Not String.IsNullOrWhiteSpace(Anterior) Then
                                        If Len(row(0)) < 9 And Len(Anterior) > Len(row(0)) Then
                                            rowIndex += 1
                                        End If
                                    End If

                                    If Len(row(0)) = 7 And Len(Anterior) = 7 Then
                                        If row(0) <> Anterior Then
                                            rowIndex += 1
                                        End If
                                    End If

                                    If Len(row(0)) = 5 And Len(Anterior) = 5 Then
                                        If row(0) <> Anterior Then
                                            rowIndex += 1
                                        End If
                                    End If

                                    If Len(row(0)) < 8 And row(0) <> Anterior Then
                                        worksheet.Cells("A" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                                    End If

                                    worksheet.Cells("A" & rowIndex).Value = row(0)
                                    worksheet.Cells("B" & rowIndex).Value = Replace(Replace(Replace(row(1), "AAA", ""), "#", " "), "ZZZ", "")
                                    worksheet.Cells("C" & rowIndex).Value = row(2)
                                    worksheet.Cells("D" & rowIndex).Value = row(3)
                                    worksheet.Cells("E" & rowIndex).Value = row(4)
                                    worksheet.Cells("F" & rowIndex).Value = row(2) + row(3) - row(4)

                                    If chkTitulos.Checked Then
                                        If Len(row(0)) = 7 Then
                                            If Len(row("Cliente")) > 4 Then
                                                Select Case row(0)
                                                    Case "1010201", "1010202", "2010101", "2010104", "2010107", "2010109", "2010111"
                                                        GConta = row(0)
                                                        GCliente = row("Cliente")
                                                        GEndCliente = row("EndCliente")
                                                        SaldoDeTitulos()

                                                        If GSaldoTitulos < 0 Then
                                                            GSaldoTitulos = GSaldoTitulos * -1
                                                        End If

                                                        Dim SR As Decimal

                                                        If worksheet.Cells("F" & rowIndex).Value < 0 Then
                                                            SR = worksheet.Cells("F" & rowIndex).Value * -1
                                                        Else
                                                            SR = worksheet.Cells("F" & rowIndex).Value
                                                        End If

                                                        worksheet.Cells("H" & rowIndex).Value = GSaldoTitulos
                                                        worksheet.Cells("I" & rowIndex).Value = SR - GSaldoTitulos
                                                End Select
                                            End If
                                        End If
                                    End If

                                    worksheet.Cells("C" & rowIndex & ":I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                    Anterior = CStr(row(0))

                                    rowIndex += 1
                                Next


                                rowIndex += 1

                                Dim dsResumo As DataSet = getDataSetResumo()

                                If dsResumo IsNot Nothing AndAlso dsResumo.Tables IsNot Nothing AndAlso dsResumo.Tables(0).Rows.Count > 0 Then
                                    For Each rowResumo As DataRow In dsResumo.Tables(0).Rows
                                        worksheet.Cells("A" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                                        worksheet.Cells("B" & rowIndex).Value = Replace(Replace(Replace(rowResumo(1), "AAA", ""), "#", " "), "ZZZ", "")
                                        worksheet.Cells("C" & rowIndex).Value = rowResumo(2)
                                        worksheet.Cells("D" & rowIndex).Value = rowResumo(3)
                                        worksheet.Cells("E" & rowIndex).Value = rowResumo(4)
                                        worksheet.Cells("F" & rowIndex).Value = rowResumo(2) + rowResumo(3) - rowResumo(4)
                                        worksheet.Cells("C" & rowIndex & ":F" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                        SdoAnterior = SdoAnterior + rowResumo(2)
                                        Debitos = Debitos + rowResumo(3)
                                        Creditos = Creditos + rowResumo(4)

                                        rowIndex += 1
                                    Next
                                End If

                                worksheet.Cells("A" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True

                                worksheet.Cells("B" & rowIndex).Value = "Total .................."
                                worksheet.Cells("C" & rowIndex).Value = SdoAnterior
                                worksheet.Cells("D" & rowIndex).Value = Debitos
                                worksheet.Cells("E" & rowIndex).Value = Creditos
                                worksheet.Cells("F" & rowIndex).Value = SdoAnterior + Debitos - Creditos
                                worksheet.Cells("C" & rowIndex & ":F" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                'setando autofit nas células da planilha
                                worksheet.Cells.AutoFitColumns(0)

                                'congelando primeira linham
                                worksheet.View.FreezePanes(9, 1)

                                worksheet.Column(1).Width = 20

                                'salvando planilha do excel
                                package.Save()
                            End Using
                        End Using

                        'download do arquivo pelo browser
                        Funcoes.AbrirExcel(Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSetResumo() As DataSet
        Dim sql As String = ""

        If rdAuxiliar.Checked Then
            sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta From (" & vbCrLf & _
                  " SELECT     LEFT(Razao.Conta_Id, 1) AS Conta,  ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo , LEFT(Razao.Conta_Id, 1) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                  "      sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                  "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                  "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos " & vbCrLf & _
                  " FROM  Razao INNER JOIN" & vbCrLf & _
                  "       PlanoDeContas ON LEFT(Razao.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
            Else
                sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            End If

            If chkCompensacao.Checked Then
                sql &= " And LEFT(Razao.Conta_Id, 3) not  in(105, 205)" & vbCrLf
            End If
            sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                   " Group  By LEFT(Razao.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf & _
                   " ) AS   Consulta " & vbCrLf & _
                   " Where  Titulo <> '' And (SaldoAnterior + Debitos + Creditos) <> 0" & vbCrLf & _
                   " Order  by Conta, Titulo" & vbCrLf

        ElseIf rdGerencial.Checked Then
            sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta From (" & vbCrLf & _
                  " SELECT     LEFT(Razao.Conta_Id, 1) AS Conta,  ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo , LEFT(Razao.Conta_Id, 1) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                  "      sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                  "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                  "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos " & vbCrLf & _
                  " FROM  Gerencial.dbo.Razao INNER JOIN" & vbCrLf & _
                  "       PlanoDeContas ON LEFT(Razao.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
            Else
                sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            End If

            If chkCompensacao.Checked Then
                sql &= " And LEFT(Razao.Conta_Id, 3) not  in(105, 205)" & vbCrLf
            End If

            sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                   " Group  By LEFT(Razao.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf & _
                   " ) AS   Consulta " & vbCrLf & _
                   " Where  Titulo <> '' And (SaldoAnterior + Debitos + Creditos) <> 0" & vbCrLf & _
                   " Order  by Conta, Titulo" & vbCrLf

        ElseIf rdAdiantamentos.Checked Then
            sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta" & vbCrLf & _
                  "  From (" & vbCrLf & _
                  "        SELECT LEFT(A.Conta_Id, 1) AS Conta," & vbCrLf & _
                  "               ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo," & vbCrLf & _
                  "               LEFT(A.Conta_Id, 1) AS Conta_Id," & vbCrLf & _
                  "               PlanoDeContas.Titulo as Nome," & vbCrLf & _
                  "               sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                  "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                  "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos " & vbCrLf & _
                  "          FROM Gerencial.dbo.vw_Adiantamento A" & vbCrLf & _
                  "         INNER JOIN PlanoDeContas" & vbCrLf & _
                  "            ON LEFT(A.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
            Else
                sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            End If

            If chkCompensacao.Checked Then
                sql &= " And LEFT(A.Conta_Id, 3) not  in(105, 205)" & vbCrLf
            End If

            sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                   " Group  By LEFT(A.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf & _
                   " ) AS   Consulta " & vbCrLf & _
                   " Where  Titulo <> '' And (SaldoAnterior + Debitos + Creditos) <> 0" & vbCrLf & _
                   " Order  by Conta, Titulo" & vbCrLf
        End If
        Return Banco.ConsultaDataSet(sql, "Resumo")
    End Function

    Private Sub SaldoDeTitulos()
        If IsNumeric(GCliente) Then
            Dim Sqla As String = " Select isnull(Sum(Apagar - AReceber), 0) as Saldo From (" & vbCrLf & _
                                 " select  Sum(ValorDoDocumento) as APagar, 0 as AReceber from ContasAPagar" & vbCrLf & _
                                 "        where Empresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                 "         And Cliente = '" & GCliente & "'" & vbCrLf & _
                                 "         And EndCliente = " & GEndCliente & vbCrLf & _
                                 "         And UsuarioInclusaoData <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                                 "         And ContaContabilCliente = '" & GConta & "'" & vbCrLf & _
                                 "         And (Provisao = 2 or (Provisao = 1 And Baixa > '" & txtDataFinal.Text.ToSqlDate() & "'))" & vbCrLf & _
                                 "         And Situacao = 1" & vbCrLf & _
                                 " Union" & vbCrLf & _
                                 " select 0 as APagar, Sum(ValorDoDocumento) as AReceber from ContasAReceber" & vbCrLf & _
                                 "        where Empresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                 "         And Cliente = '" & GCliente & "'" & vbCrLf & _
                                 "         And EndCliente = " & GEndCliente & vbCrLf & _
                                 "         And UsuarioInclusaoData <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                                 "         And ContaContabilCliente = '" & GConta & "'" & vbCrLf & _
                                 "         And (Provisao = 2 or (Provisao = 1 And Baixa > '" & txtDataFinal.Text.ToSqlDate() & "'))" & vbCrLf & _
                                 "         And Situacao = 1" & vbCrLf & _
                                 "         ) as Consulta" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sqla, "Saldo")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    GSaldoTitulos = row("Saldo")
                Next
            End If
        End If
    End Sub

    Private Function getDataSetCompleto() As DataSet
        Try
            Dim sql As String = ""
            If rdAuxiliar.Checked Then
                sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta, Cliente, EndCliente From (" & vbCrLf & _
                      " SELECT     LEFT(Razao.Conta_Id, 1) AS Conta,  ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo , LEFT(Razao.Conta_Id, 1) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                      "      sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                      "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                      "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                      "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                      " FROM  Razao INNER JOIN" & vbCrLf & _
                      "       PlanoDeContas ON LEFT(Razao.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                If chkCompensacao.Checked = True Then
                    sql &= " And LEFT(Razao.Conta_Id, 3) not  in(105, 205)" & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf

                '-----Nivel 2 --------------------
                sql &= " Union" & vbCrLf & _
                       " SELECT LEFT(Razao.Conta_Id, 3) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(Razao.Conta_Id, 3) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM   Razao INNER JOIN" & vbCrLf & _
                       "        PlanoDeContas ON LEFT(Razao.Conta_Id, 3) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 3), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 3 ----------------------
                sql &= " Union" & vbCrLf & _
                       " SELECT LEFT(Razao.Conta_Id, 5) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(Razao.Conta_Id, 5) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM   Razao INNER JOIN" & vbCrLf & _
                       "        PlanoDeContas ON LEFT(Razao.Conta_Id, 5) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 5), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 4 --------------------
                sql &= "    Union" & vbCrLf & _
                       " SELECT LEFT(Razao.Conta_Id, 7) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(Razao.Conta_Id, 7) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM   Razao INNER JOIN" & vbCrLf & _
                       "        PlanoDeContas ON LEFT(Razao.Conta_Id, 7) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 7), PlanoDeContas.Titulo" & vbCrLf

                If chkClientes.Checked Then

                    '---- Nivel 5 Clientes ------------------
                    sql &= " Union" & vbCrLf & _
                           " SELECT Razao.Conta_Id as Conta, ISNULL('ZZZ' + LEFT(Clientes.Nome, 35) + ' - ' + LEFT(Clientes.Cidade, 18) + ' |' + Razao.Cliente_Id + '|' + CONVERT(varchar, Razao.EndCliente_Id)," & vbCrLf & _
                           "        'ZZZ' + Razao.Cliente_Id + '-' + CONVERT(varchar, Razao.EndCliente_Id)) AS Titulo, Razao.Conta_Id, Clientes.Nome," & vbCrLf & _
                           "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                           "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                           "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                           "        Razao.Cliente_Id as Cliente, Razao.EndCliente_Id as EndCliente" & vbCrLf & _
                           " FROM  Razao LEFT OUTER JOIN" & vbCrLf & _
                           "       Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id AND Razao.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf

                    If chkConsolidarEmpresa.Checked Then
                        sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                    Else
                        sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                    End If

                    sql &= " AND (LEN(Razao.Conta_Id) = 7)" & vbCrLf & _
                           " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                           " GROUP BY Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, Clientes.Nome, Clientes.Cidade" & vbCrLf
                End If

                '-------- Nivel 5 Analiticas -------------------
                sql &= "    Union" & vbCrLf & _
                       " SELECT Razao.Conta_Id AS Conta, 'AAA' + PlanoDeContas.Titulo + '# ' AS Titulo, Razao.Conta_Id, 'a' as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM  Razao LEFT OUTER JOIN PlanoDeContas ON Razao.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " AND (LEN(Razao.Conta_Id) = 9)" & vbCrLf & _
                       " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " GROUP BY Razao.Conta_Id, PlanoDeContas.Titulo" & vbCrLf & _
                       " ) AS   Consulta " & vbCrLf & _
                       " Where  Titulo <> '' And (SaldoAnterior + Debitos + Creditos) <> 0" & vbCrLf & _
                       " Order  by Conta, Titulo" & vbCrLf

            ElseIf rdGerencial.Checked Then
                '-----------Balancete --------------------------------------
                sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta, Cliente, EndCliente From (" & vbCrLf & _
                      " SELECT     LEFT(Razao.Conta_Id, 1) AS Conta,  ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo , LEFT(Razao.Conta_Id, 1) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                      "      sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                      "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                      "      sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                      "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                      " FROM  Gerencial.dbo.Razao INNER JOIN" & vbCrLf & _
                      "       Insol.dbo.PlanoDeContas ON LEFT(Razao.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                If chkCompensacao.Checked Then
                    sql &= " And LEFT(Razao.Conta_Id, 3) not  in(105, 205)" & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf

                '-----Nivel 2 --------------------
                sql &= " Union" & vbCrLf & _
                       " SELECT LEFT(Razao.Conta_Id, 3) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(Razao.Conta_Id, 3) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM   Gerencial.dbo.Razao INNER JOIN" & vbCrLf & _
                       "        Insol.dbo.PlanoDeContas ON LEFT(Razao.Conta_Id, 3) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 3), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 3 ----------------------

                sql &= " Union" & vbCrLf & _
                       " SELECT LEFT(Razao.Conta_Id, 5) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(Razao.Conta_Id, 5) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM   Gerencial.dbo.Razao INNER JOIN" & vbCrLf & _
                       "        Insol.dbo.PlanoDeContas ON LEFT(Razao.Conta_Id, 5) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 5), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 4 --------------------

                sql &= "    Union" & vbCrLf & _
                       " SELECT LEFT(Razao.Conta_Id, 7) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(Razao.Conta_Id, 7) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM   Gerencial.dbo.Razao INNER JOIN" & vbCrLf & _
                       "        Insol.dbo.PlanoDeContas ON LEFT(Razao.Conta_Id, 7) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(Razao.Conta_Id, 7), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 5 Clientes ------------------

                If chkClientes.Checked Then
                    sql &= " Union" & vbCrLf & _
                           " SELECT Razao.Conta_Id as Conta, ISNULL('ZZZ' + LEFT(Clientes.Nome, 35) + ' - ' + LEFT(Clientes.Cidade, 18) + ' |' + Razao.Cliente_Id + '|' + CONVERT(varchar, Razao.EndCliente_Id)," & vbCrLf & _
                           "        'ZZZ' + Razao.Cliente_Id + '-' + CONVERT(varchar, Razao.EndCliente_Id)) AS Titulo, Razao.Conta_Id, Clientes.Nome," & vbCrLf & _
                           "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                           "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                           "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                           "        Razao.Cliente_Id as Cliente, Razao.EndCliente_Id as EndCliente" & vbCrLf & _
                           " FROM  Gerencial.dbo.Razao LEFT OUTER JOIN" & vbCrLf & _
                           "       Insol.dbo.Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id AND Razao.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf

                    If chkConsolidarEmpresa.Checked Then
                        sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                    Else
                        sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                    End If

                    sql &= " AND (LEN(Razao.Conta_Id) = 7)" & vbCrLf & _
                           " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                           " GROUP BY Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, Clientes.Nome, Clientes.Cidade" & vbCrLf
                End If

                '-------- Nivel 5 Analiticas -------------------

                sql &= "    Union" & vbCrLf & _
                       " SELECT Razao.Conta_Id AS Conta, 'AAA' + PlanoDeContas.Titulo + '# ' AS Titulo, Razao.Conta_Id, 'a' as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "      '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       " FROM  Gerencial.dbo.Razao LEFT OUTER JOIN Insol.dbo.PlanoDeContas ON Razao.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(Razao.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " AND (LEN(Razao.Conta_Id) = 9)" & vbCrLf & _
                       " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " GROUP BY Razao.Conta_Id, PlanoDeContas.Titulo" & vbCrLf & _
                       " ) AS   Consulta " & vbCrLf & _
                       " Where  Titulo <> '' And (SaldoAnterior + Debitos + Creditos) <> 0" & vbCrLf & _
                       " Order  by Conta, Titulo" & vbCrLf
            ElseIf rdAdiantamentos.Checked Then
                '-----------Balancete --------------------------------------
                sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta, Cliente, EndCliente" & vbCrLf & _
                      "  From (" & vbCrLf & _
                      "        SELECT LEFT(A.Conta_Id, 1) AS Conta,  ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo , LEFT(A.Conta_Id, 1) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                      "               sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                      "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                      "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                      "               '' as Cliente, 0 as EndCliente" & vbCrLf & _
                      "          FROM Gerencial.dbo.VW_Adiantamento A" & vbCrLf & _
                      "         INNER JOIN Insol.dbo.PlanoDeContas" & vbCrLf & _
                      "            ON LEFT(A.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                If chkCompensacao.Checked Then
                    sql &= " And LEFT(A.Conta_Id, 3) not  in(105, 205)" & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(A.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf

                '-----Nivel 2 --------------------
                sql &= " Union" & vbCrLf & _
                       " SELECT LEFT(A.Conta_Id, 3) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(A.Conta_Id, 3) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "        '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       "   FROM Gerencial.dbo.vw_Adiantamento A" & vbCrLf & _
                       "  INNER JOIN Insol.dbo.PlanoDeContas" & vbCrLf & _
                       "     ON LEFT(A.Conta_Id, 3) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(A.Conta_Id, 3), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 3 ----------------------

                sql &= " Union" & vbCrLf & _
                       " SELECT LEFT(A.Conta_Id, 5) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(A.Conta_Id, 5) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "        '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       "   FROM Gerencial.dbo.VW_Adiantamento A" & vbCrLf & _
                       "  INNER JOIN Insol.dbo.PlanoDeContas" & vbCrLf & _
                       "     ON LEFT(A.Conta_Id, 5) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group  By LEFT(A.Conta_Id, 5), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 4 --------------------

                sql &= "    Union" & vbCrLf & _
                       " SELECT LEFT(A.Conta_Id, 7) AS Conta,  ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo, LEFT(A.Conta_Id, 7) AS Conta_Id, PlanoDeContas.Titulo as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "        '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       "   FROM Gerencial.dbo.vw_Adiantamento A" & vbCrLf & _
                       "  INNER JOIN Insol.dbo.PlanoDeContas" & vbCrLf & _
                       "     ON LEFT(A.Conta_Id, 7) = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= "   And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " Group By LEFT(A.Conta_Id, 7), PlanoDeContas.Titulo" & vbCrLf

                '---- Nivel 5 Clientes ------------------

                If chkClientes.Checked Then
                    sql &= " Union" & vbCrLf & _
                           " SELECT A.Conta_Id as Conta, ISNULL('ZZZ' + LEFT(Clientes.Nome, 35) + ' - ' + LEFT(Clientes.Cidade, 18) + ' |' + A.Cliente_Id + '|' + CONVERT(varchar, A.EndCliente_Id)," & vbCrLf & _
                           "        'ZZZ' + A.Cliente_Id + '-' + CONVERT(varchar, A.EndCliente_Id)) AS Titulo, A.Conta_Id, Clientes.Nome," & vbCrLf & _
                           "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                           "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                           "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                           "        A.Cliente_Id as Cliente, A.EndCliente_Id as EndCliente" & vbCrLf & _
                           "   FROM Gerencial.dbo.vw_Adiantamento A" & vbCrLf & _
                           "   LEFT JOIN Insol.dbo.Clientes" & vbCrLf & _
                           "     ON A.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf & _
                           "    AND A.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf

                    If chkConsolidarEmpresa.Checked Then
                        sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                    Else
                        sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                    End If

                    sql &= " AND (LEN(A.Conta_Id) = 7)" & vbCrLf & _
                           " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                           " GROUP BY A.Conta_Id, A.Cliente_Id, A.EndCliente_Id, Clientes.Nome, Clientes.Cidade" & vbCrLf
                End If

                '-------- Nivel 5 Analiticas -------------------

                sql &= "    Union" & vbCrLf & _
                       " SELECT A.Conta_Id AS Conta, 'AAA' + PlanoDeContas.Titulo + '# ' AS Titulo, A.Conta_Id, 'a' as Nome," & vbCrLf & _
                       "        sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "        sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "        '' as Cliente, 0 as EndCliente" & vbCrLf & _
                       "   FROM Gerencial.dbo.vw_Adiantamento A" & vbCrLf & _
                       "   LEFT JOIN Insol.dbo.PlanoDeContas" & vbCrLf & _
                       "     ON A.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= " Where (SUBSTRING(A.Empresa_Id, 1, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & " ')" & vbCrLf
                Else
                    sql &= " Where A.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And A.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                sql &= " AND (LEN(A.Conta_Id) = 9)" & vbCrLf & _
                       " And Not (year(movimento_Id) = " & Year(CDate(txtDataInicial.Text)) & " and Lote_Id  In (7500, 7600)) " & vbCrLf & _
                       " GROUP BY A.Conta_Id, PlanoDeContas.Titulo" & vbCrLf & _
                       " ) AS   Consulta " & vbCrLf & _
                       " Where  Titulo <> '' And (SaldoAnterior + Debitos + Creditos) <> 0" & vbCrLf & _
                       " Order  by Conta, Titulo" & vbCrLf
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Consulta")

            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Sub limpar()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)

        txtDataInicial.Text = DateSerial(Today.Year, Today.Month, 1)
        txtDataFinal.Text = DateSerial(Today.Year, Today.Month + 1, 0)

        'chkConsolidarEmpresa.Checked = False
        'chkCompensacao.Checked = True
        'chkIsolar.Checked = True
        'chkTitulos.Checked = False
        'chkGeral.Checked = False   
        rdAuxiliar.Checked = True
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub SaldoDeTitulos(ByVal Empresa As String, ByVal Cliente As String, ByVal Conta As String, ByRef Saldo As Decimal)
        Try
            If IsNumeric(Cliente.Split("-")(0)) Then
                Dim sql As String = ""
                If FinanceiroNovo Then
                    sql &= "   Select isnull(Sum(Apagar - AReceber), 0) as Saldo From (                         " & vbCrLf & _
                            "   Select SUM(tc.ValorOficial) as APagar, 0 as AReceber                             " & vbCrLf & _
                            "      from Titulos t                                                                " & vbCrLf & _
                            "     Inner Join TitulosxContaContabil tc                                            " & vbCrLf & _
                            "        on tc.Titulo_Id = t.Titulo_Id                                               " & vbCrLf & _
                            "   	and tc.Conta_Id = t.ContaContabilCliFor                                      " & vbCrLf & _
                            "   Where t.RecPag = 'P'                                                             " & vbCrLf & _
                            "   	And t.Empresa = '" & Empresa & "'" & vbCrLf & _
                            "   	And t.CliFor = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                            "   	And t.EnderecoCliFor = " & Cliente.Split("-")(1) & vbCrLf & _
                            "   	and t.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                            "   	And T.ContaContabilCliFor = '" & Conta & "'" & vbCrLf & _
                            "   	And (t.Provisao = 3 or (t.Provisao = 1 And t.DataBaixa > '" & txtDataFinal.Text.ToSqlDate() & "'))      " & vbCrLf & _
                            "   	And t.Situacao = 1                                                           " & vbCrLf & _
                            "   UNION ALL                                                                        " & vbCrLf & _
                            "   Select 0 as APagar, SUM(tc.ValorOficial) as AReceber                             " & vbCrLf & _
                            "      from Titulos t                                                                " & vbCrLf & _
                            "     Inner Join TitulosxContaContabil tc                                            " & vbCrLf & _
                            "        on tc.Titulo_Id = t.Titulo_Id                                               " & vbCrLf & _
                            "   	and tc.Conta_Id = t.ContaContabilCliFor                                      " & vbCrLf & _
                            "   Where t.RecPag = 'R'" & vbCrLf & _
                            "   	And t.Empresa = '" & Empresa & "'" & vbCrLf & _
                            "   	And t.CliFor = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                            "   	And t.EnderecoCliFor = " & Cliente.Split("-")(1) & vbCrLf & _
                            "   	and t.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                            "   	And T.ContaContabilCliFor = '" & Conta & "'" & vbCrLf & _
                            "   	And (t.Provisao = 3 or (t.Provisao = 1 And t.DataBaixa > '" & txtDataFinal.Text.ToSqlDate() & "'))      " & vbCrLf & _
                            "   	And t.Situacao = 1" & vbCrLf & _
                            "   ) as Consulta	      " & vbCrLf
                Else
                    sql = " Select isnull(Sum(Apagar - AReceber), 0) as Saldo From (" & vbCrLf & _
                          " select  Sum(ValorDoDocumento) as APagar, 0 as AReceber from ContasAPagar" & vbCrLf & _
                          "        where Empresa = '" & Empresa & "'" & vbCrLf & _
                          "         And Cliente = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                          "         And EndCliente = " & Cliente.Split("-")(1) & vbCrLf & _
                          "         And UsuarioInclusaoData <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                          "         And ContaContabilCliente = '" & Conta & "'" & vbCrLf & _
                          "         And (Provisao = 2 or (Provisao = 1 And Baixa > '" & txtDataFinal.Text.ToSqlDate() & "'))" & vbCrLf & _
                          "         And Situacao = 1" & vbCrLf & _
                          " Union" & vbCrLf & _
                          " select 0 as APagar, Sum(ValorDoDocumento) as AReceber from ContasAReceber" & vbCrLf & _
                          "        where Empresa = '" & Empresa & "'" & vbCrLf & _
                          "         And Cliente = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                          "         And EndCliente = " & Cliente.Split("-")(1) & vbCrLf & _
                          "         And UsuarioInclusaoData <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                          "         And ContaContabilCliente = '" & Conta & "'" & vbCrLf & _
                          "         And (Provisao = 2 or (Provisao = 1 And Baixa > '" & txtDataFinal.Text.ToSqlDate() & "'))" & vbCrLf & _
                          "         And Situacao = 1" & vbCrLf & _
                          "         ) as Consulta" & vbCrLf
                End If

                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "SaldoTitulo")

                For Each row As DataRow In ds.Tables(0).Rows
                    Saldo = row("Saldo")
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "BalanceteAuxiliarVba")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class