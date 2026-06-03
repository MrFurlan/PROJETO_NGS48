Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PosicaoDePedido
    Inherits BasePage

    Dim Sql As String = ""
    Dim SqlArray As New ArrayList
    Dim i As Integer = 0
    Dim PeriodoInicial As String = ""
    Dim PeriodoFinal As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PosicaoDePedido", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("PlanilhasExcel")) Then
                        Directory.CreateDirectory(Server.MapPath("PlanilhasExcel"))
                    End If
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    CargaUnidade()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function Validar() As Boolean
        If ddlUnidade.SelectedIndex < 1 Then
            MsgBox(Me.Page, "Informe a unidade de negócio.")
            Return False
        End If
        If DdlEmpresa.SelectedIndex < 1 Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe o período inicial")
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período final.")
            Return False
        End If

        Return True
    End Function

    Private Sub Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getDataSet() As DataSet
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy-MM-dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy-MM-dd")

        AtualizaPosicaoDePedidos()

        Sql &= "    SELECT Emp.Reduzido AS RedEmpresa,                                       " & vbCrLf &
               "          Emp.Cliente_Id AS CNPJEmpresa,                                   " & vbCrLf &
               "          --Emp.Endereco_Id AS EndEmpresa,                                 " & vbCrLf &
               "          Emp.Nome AS NomeEmpresa,                                         " & vbCrLf &
               "   	   Emp.Cidade AS CidadeEmpresa,                                     " & vbCrLf &
               "   	   --Emp.Estado AS UFEmpresa,                                       " & vbCrLf &
               "   	  cliDeposito.Cliente_Id AS CNPJDeposito,                           " & vbCrLf &
               "   	   --cli.Endereco_Id AS EndCliente,                                 " & vbCrLf &
               "   	   cliDeposito.Nome AS NomeDeposito,                                " & vbCrLf &
               "   	   cliDeposito.Cidade AS CidadeDeposito,                            " & vbCrLf &
               "   	   --cli.Estado AS UFCliente,                                       " & vbCrLf &
               "   	   cli.Cliente_Id as Cliente,                                       " & vbCrLf &
               "   	   cli.Endereco_Id as EndCliente,                                   " & vbCrLf &
               "   	   cli.Nome as NomeCliente,                                         " & vbCrLf &
               "   	   cli.Cidade as CidadeCliente,                                     " & vbCrLf &
               "   	   sOp.Operacao_Id AS Operacao,                                     " & vbCrLf &
               "   	   sOp.SubOperacoes_Id AS SubOperacao,                              " & vbCrLf &
               "   	   sOp.Descricao AS NomeOperacao,                                   " & vbCrLf &
               "   	   p.Moeda,                                                         " & vbCrLf &
               "   	   m.Descricao AS NomeMoeda,                                        " & vbCrLf &
               "   	   p.Indexador,                                                     " & vbCrLf &
               "   	   ind.Descricao AS NomeIndexador,                                  " & vbCrLf &
               "   	   p.Finalidade,                                                    " & vbCrLf &
               "   	   fin.Descricao AS NomeFinalidade,                                 " & vbCrLf &
               "   	   p.FreteCIFFOB,                                                   " & vbCrLf &
               "   	   p.CondicaoPagamento,                                             " & vbCrLf &
               "   	   pgto.Descricao AS NomeCondicaoDePagamento,                       " & vbCrLf &
               "   	   p.DataPedido,                                                    " & vbCrLf &
               "   	   p.DataEntrega,                                                   " & vbCrLf &
               "   	   #Posicao.Pedido_Id AS Pedido,                                    " & vbCrLf &
               "   	   #Posicao.Produto,                                                " & vbCrLf &
               "   	   prd.Nome AS NomeDoProduto,                                       " & vbCrLf &
               "   	   #Posicao.Contratado,                                             " & vbCrLf &
               "   	   #Posicao.UnitarioOficial,                                        " & vbCrLf &
               "   	   isnull(#Posicao.TotalOficial, 0) as TotalOficial,                " & vbCrLf &
               "   	   #Posicao.UnitarioMoeda,                                          " & vbCrLf &
               "   	   #Posicao.TotalMoeda,                                             " & vbCrLf &
               "   	   #Posicao.Laudo As Laudo,                                         " & vbCrLf &
               "   	   isnull(#Posicao.Entregue, 0) AS PesoFiscal,                      " & vbCrLf &
               "   	   isnull(#Posicao.AEntregar, 0) as AEntregar,                      " & vbCrLf &
               "   	   isnull(#Posicao.QuantidadeFixado, 0) AS Fixado,                  " & vbCrLf &
               "   	   isnull(#Posicao.AFixar, 0) as AFixar,                            " & vbCrLf &
               "   	   isnull(#Posicao.Pago, 0) as Pago,                                " & vbCrLf &
               "   	   isnull(#Posicao.PagoNaoRecebido, 0) AS PagoNaoEntregue,          " & vbCrLf &
               "   	   isnull(#Posicao.RecebidoNaoPago, 0) AS EntregueNaoPago,          " & vbCrLf &
               "   	   isnull(#Posicao.Adiantamento, 0) as Adiantamento,                " & vbCrLf &
               "   	   p.UsuarioInclusao,                                               " & vbCrLf &
               "   	   p.UsuarioAlteracao,                                              " & vbCrLf &
               "   	   p.UsuarioLiberacao,                                              " & vbCrLf &
               "   	   --Antigo Cortava Aqui                                            " & vbCrLf &
               "   	   #Posicao.Programado_PagarOficial,                                " & vbCrLf &
               "   	   #Posicao.Baixado_PagarOficial,                                   " & vbCrLf &
               "   	   #Posicao.Programado_PagarMoeda,                                  " & vbCrLf &
               "   	   #Posicao.Baixado_PagarMoeda,                                     " & vbCrLf &
               "   	   #Posicao.Programado_ReceberOficial,                              " & vbCrLf &
               "   	   #Posicao.Programado_ReceberMoeda,                                " & vbCrLf &
               "   	   #Posicao.Baixado_ReceberOficial,                                 " & vbCrLf &
               "   	   #Posicao.Baixado_ReceberMoeda,                                   " & vbCrLf &
               "   	   #Posicao.Programado_SaldoOficial,                                " & vbCrLf &
               "   	   #Posicao.Baixado_SaldoOficial,                                   " & vbCrLf &
               "   	   #Posicao.Programado_SaldoMoeda,                                  " & vbCrLf &
               "   	   #Posicao.Baixado_SaldoMoeda,                                     " & vbCrLf &
               "   	   #Posicao.OrigemDestino,                                          " & vbCrLf &
               "   	   #Posicao.EntradaSaida,                                           " & vbCrLf &
               "   	   #Posicao.MoedaPedido,                                            " & vbCrLf &
               "   	   #Posicao.ValorFixadoOficial ,                                    " & vbCrLf &
               "   	   #Posicao.ValorFixadoMoeda,                                       " & vbCrLf &
               "   	   IsNull(#Posicao.ValorLiquidoDeNota,0) as ValorLiquidoDeNota,     " & vbCrLf &
               "   	   isnull(#Posicao.Complementacoes, 0) as Complementacoes,          " & vbCrLf &
               "          isnull(#Posicao.TotalLiquidoOficial, 0) as TotalLiquidoOficial   " & vbCrLf &
               "     FROM #Posicao                                                         " & vbCrLf &
               "    INNER JOIN Pedidos p                                                   " & vbCrLf &
               "   	ON p.Empresa_Id    = #Posicao.Empresa_Id                            " & vbCrLf &
               "      AND p.EndEmpresa_Id = #Posicao.EndEmpresa_Id                         " & vbCrLf &
               "      AND p.Pedido_Id     = #Posicao.Pedido_Id                             " & vbCrLf &
               "      Inner Join PedidosXDepositos pxd                                     " & vbCrLf &
               "   	ON p.Empresa_Id = pxd.Empresa_Id                                    " & vbCrLf &
               "   				AND p.EndEmpresa_Id = pxd.EndEmpresa_Id                 " & vbCrLf &
               "   				AND p.Pedido_Id = pxd.Pedido_Id                         " & vbCrLf &
               "   				And pxd.Tipo = 'DE'                                     " & vbCrLf &
               "   				And Principal = 1                                       " & vbCrLf &
               "   	Inner join Clientes cliDeposito                                     " & vbCrLf &
               "   		on cliDeposito.Cliente_Id = pxd.Deposito_Id                     " & vbCrLf &
               "   		and cliDeposito.Endereco_Id = pxd.EndDeposito_Id                " & vbCrLf &
               "    INNER JOIN Clientes cli                                                " & vbCrLf &
               "   	ON cli.Cliente_Id  = p.Cliente                                      " & vbCrLf &
               "      AND cli.Endereco_Id = p.EndCliente                                   " & vbCrLf &
               "     LEFT JOIN Indexadores ind                                             " & vbCrLf &
               "   	ON ind.Indexador_Id = p.Indexador                                   " & vbCrLf &
               "     LEFT JOIN Pagamentos pgto                                             " & vbCrLf &
               "   	ON pgto.Pagamento_Id = p.CondicaoPagamento                          " & vbCrLf &
               "     LEFT JOIN Moedas m                                                    " & vbCrLf &
               "   	ON m.Moeda_Id = p.Moeda                                             " & vbCrLf &
               "     LEFT JOIN Finalidades fin                                             " & vbCrLf &
               "   	ON fin.Finalidade_Id = p.Finalidade                                 " & vbCrLf &
               "     LEFT JOIN SubOperacoes sOp                                            " & vbCrLf &
               "   	ON sOp.Operacao_Id     = p.Operacao                                 " & vbCrLf &
               "      AND sOp.SubOperacoes_Id = p.SubOperacao                              " & vbCrLf &
               "     LEFT JOIN Clientes AS Emp                                             " & vbCrLf &
               "   	ON Emp.Cliente_Id  = p.Empresa_Id                                   " & vbCrLf &
               "      AND Emp.Endereco_Id = p.EndEmpresa_Id                                " & vbCrLf &
               "     LEFT JOIN ComprasXProdutos cXp                                        " & vbCrLf &
               "   	ON cXp.Produto_Id = p.Carteira                                      " & vbCrLf &
               "     LEFT JOIN Produtos prd                                                " & vbCrLf &
               "       ON prd.Produto_Id = #Posicao.Produto                                " & vbCrLf &
               "    WHERE prd.Agrupar = 'N'                 " & vbCrLf

        If Not String.IsNullOrEmpty(DdlEmpresa.SelectedValue) Then
            If cbConsolidado.Checked Then
                Sql &= "And Left(p.Empresa_Id, 8) = " & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & vbCrLf
            Else
                Sql &= "       And p.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                       "       And p.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            Sql &= "And p.DataPedido between '" & CDate(txtDataInicial.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        Sql &= "   order by #Posicao.Pedido_Id                                             " & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "PosicaoDePedidos")
    End Function

    Private Sub Processar()
        Try
            If rdoPlanilha.Checked Then
                Dim ds As DataSet = getDataSet()

                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    Throw New Exception("Nenhum resultado encontrado.")
                End If

                Dim param As New Dictionary(Of String, eTipoCampo)
                param.Add("DataPedido", eTipoCampo.Data)
                param.Add("DataEntrega", eTipoCampo.Data)
                param.Add("UnitarioOficial", eTipoCampo.ValorSemTotalizador)
                param.Add("TotalOficial", eTipoCampo.ValorSemTotalizador)
                param.Add("UnitarioMoeda", eTipoCampo.ValorSemTotalizador)
                param.Add("TotalMoeda", eTipoCampo.ValorSemTotalizador)
                param.Add("Contratado", eTipoCampo.Numerico)
                param.Add("Entregue", eTipoCampo.Numerico)
                param.Add("AEntregar", eTipoCampo.Numerico)
                param.Add("Fixado", eTipoCampo.Numerico)
                param.Add("AFixar", eTipoCampo.Numerico)
                param.Add("Pago", eTipoCampo.Numerico)
                param.Add("PagoNaoEntregue", eTipoCampo.Numerico)
                param.Add("EntregueNaoPago", eTipoCampo.Numerico)

                Funcoes.BindExcelOffice(Me.Page, ds, "Posição de Pedidos", param)
            Else
                Dim QuebraEmpresa As String = ""
                Dim QuebraEntradaSaida As String = ""
                Dim QuebraProduto As String = ""
                Dim Anterior As String = ""
                Dim Contratado As Double
                Dim Entregue As Double
                Dim ValorTotal As Double
                Dim ValorPago As Double
                Dim Pago As Double
                Dim APagar As Double
                Dim PesoFiscal As Double
                Dim ValorFiscal As Double
                Dim PesoFiscal_Dif As Double
                Dim ValorFiscal_Dif As Double
                Dim PesoFinanceiro As Double
                Dim ValorFinanceiro As Double
                Dim PesoFinanceiro_Dif As Double
                Dim ValorFinanceiro_Dif As Double

                Dim ds As DataSet = getDataSetExcel()

                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    Throw New Exception("Nenhum resultado encontrado.")
                End If

                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando planilha Posição de Pedidos
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Posição de Pedido")

                        'criando linha com o cabeçalho da planilha
                        Dim rowIndex As Integer = 1
                        Dim columnIndex As Integer = 1

                        For Each dr As DataRow In ds.Tables(0).Rows
                            If rowIndex = 1 Then
                                For i As Integer = 0 To 5
                                    If rowIndex <> 4 Then
                                        Using range = worksheet.Cells("A" & rowIndex & ":M" & rowIndex)
                                            range("A" & rowIndex & ":M" & rowIndex).Merge = True
                                            range.Style.Font.Bold = True
                                            range.Style.Font.Name = "Calibri"
                                            If rowIndex > 1 Then
                                                range.Style.Font.Size = 11
                                            Else
                                                range.Style.Font.Size = 12
                                            End If
                                        End Using
                                    End If
                                    rowIndex += 1
                                Next
                                worksheet.Cells("A7:F7").Merge = True
                                worksheet.Cells("A7:F7").Style.Font.Bold = True
                                worksheet.Cells("L7:N7").Merge = True
                                worksheet.Cells("L7:N7").Style.Font.Bold = True
                                worksheet.Cells("S7:V7").Merge = True
                                worksheet.Cells("S7:V7").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                                worksheet.Cells("X7:AB7").Merge = True
                                worksheet.Cells("X7:AB7").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                                worksheet.Cells("A1").Value = dr("NomeEmpresa")
                                worksheet.Cells("A2").Value = "           " & dr("CidadeEmpresa") & " - " & dr("UFEmpresa")
                                worksheet.Cells("A3").Value = "           CNPJ - " & Funcoes.FormatarCpfCnpj(dr("CNPJEmpresa"))
                                worksheet.Cells("A5").Value = "           Posição De Pedidos Pendentes"
                                worksheet.Cells("A7").Value = "Referente Saldos até o dia " & Now
                                worksheet.Cells("S7").Value = "Fiscal"
                                worksheet.Cells("X7").Value = "Financeiro"

                                Dim arr() As String = {"Descrição", "", "Cidade", "UF", "Pedido", "DT Pedido", "OP", "SO", "Contratado", "Entregue", "A Entregar", "Unitário", "ValorTotal", "Qt.Paga", "Qt.APagar",
                                   "ValorPago", "ValorAPagar", "", "Peso", "Valor", "Dif.Peso", "Dif.Valor", "", "Peso", "Valor", "Dif.Peso", "Dif.Valor", "Adto"}

                                For i As Integer = 0 To arr.Length - 1
                                    worksheet.Cells(8, i + 1).Value = arr(i)
                                Next

                                Using range = worksheet.Cells("A8:M8")
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                                Using range = worksheet.Cells("N8:Q8")
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 153))
                                End Using
                                Using range = worksheet.Cells("S8:V8")
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 204))
                                End Using
                                Using range = worksheet.Cells("X8:AB8")
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 255))
                                End Using
                                i = 9
                            End If
                            If dr("Empresa") <> QuebraEmpresa Then
                                If QuebraEmpresa <> "" Then
                                    i = i + 1
                                End If

                                worksheet.Cells("A7")("A" & i & ":H" & i).Merge = True
                                worksheet.Cells("A" & i & ":H" & i).Style.Font.Bold = True
                                worksheet.Cells("A" & i).Value = dr("Empresa") & "-" & dr("CidadeEmpresa") & "-" & dr("UFEmpresa")
                                QuebraEmpresa = dr("Empresa")
                                i = i + 1
                            End If

                            If dr("EntradaSaida") <> QuebraEntradaSaida Then
                                worksheet.Cells("A" & i & ":H" & i).Merge = True
                                If QuebraEntradaSaida <> "" Then

                                    '"******1*******"
                                    worksheet.Cells("I" & i).Value = Contratado
                                    worksheet.Cells("J" & i).Value = Entregue
                                    worksheet.Cells("K" & i).Value = Contratado - Entregue

                                    worksheet.Cells("M" & i).Value = ValorTotal
                                    worksheet.Cells("N" & i).Value = Pago
                                    worksheet.Cells("O" & i).Value = APagar

                                    worksheet.Cells("P" & i).Value = ValorPago

                                    'worksheet.Cells(String.Format("Q{0}", i)).Formula = String.Format("=SUM(Q12:Q{0})", i - 1)

                                    worksheet.Cells("Q" & i).Value = ValorTotal - ValorPago

                                    worksheet.Cells("S" & i).Value = PesoFiscal
                                    worksheet.Cells("T" & i).Value = ValorFiscal
                                    worksheet.Cells("U" & i).Value = PesoFiscal_Dif
                                    worksheet.Cells("V" & i).Value = ValorFiscal_Dif

                                    worksheet.Cells("X" & i).Value = PesoFinanceiro
                                    worksheet.Cells("Y" & i).Value = ValorFinanceiro
                                    worksheet.Cells("Z" & i).Value = PesoFinanceiro_Dif
                                    worksheet.Cells("AA" & i).Value = ValorFinanceiro_Dif

                                    worksheet.Cells(String.Format("I{0}:K{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("L{0}:M{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("N{0}:O{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("P{0}:Q{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("S{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("T{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("U{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("V{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("X{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("Y{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("Z{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("AA{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("I" & i & ":AA" & i).Style.Font.Bold = True

                                    Contratado = 0
                                    Entregue = 0
                                    Pago = 0
                                    APagar = 0

                                    ValorTotal = 0
                                    ValorPago = 0

                                    PesoFiscal = 0
                                    ValorFiscal = 0
                                    PesoFiscal_Dif = 0
                                    ValorFiscal_Dif = 0

                                    PesoFinanceiro = 0
                                    ValorFinanceiro = 0
                                    PesoFinanceiro_Dif = 0
                                    ValorFinanceiro_Dif = 0

                                    Using range = worksheet.Cells("I" & i & ":I" & i)
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Left.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                                    End Using

                                    Using range = worksheet.Cells("J" & i & ":AA" & i)
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                                    End Using

                                    Using range = worksheet.Cells("AB" & i & ":AB" & i)
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Right.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                                    End Using
                                    i = i + 1
                                End If

                                If dr("EntradaSaida") = "E" Then
                                    worksheet.Cells("A" & i).Value = "   Compras"
                                Else
                                    worksheet.Cells("A" & i).Value = "   Vendas"
                                End If
                                worksheet.Cells("A" & i & ":H" & i).Style.Font.Bold = True
                                QuebraEntradaSaida = dr("EntradaSaida")
                                i = i + 1

                                worksheet.Cells("A" & i & ":H" & i).Merge = True
                                worksheet.Cells("A" & i).Value = "      " & dr("Produto") & "-" & dr("NomeDoProduto")
                                worksheet.Cells("A" & i & ":H" & i).Style.Font.Bold = True
                                QuebraProduto = dr("Produto")
                                i = i + 1
                            End If

                            If dr("Produto") <> QuebraProduto Then
                                If QuebraProduto <> "" Then

                                    '******2*******"
                                    worksheet.Cells("I" & i).Value = Contratado
                                    worksheet.Cells("J" & i).Value = Entregue
                                    worksheet.Cells("K" & i).Value = Contratado - Entregue
                                    worksheet.Cells("K" & i).Value = Contratado - Entregue

                                    worksheet.Cells("M" & i).Value = ValorTotal
                                    worksheet.Cells("N" & i).Value = Pago
                                    worksheet.Cells("O" & i).Value = APagar

                                    worksheet.Cells("P" & i).Value = ValorPago
                                    worksheet.Cells("Q" & i).Value = ValorTotal - ValorPago

                                    worksheet.Cells("S" & i).Value = PesoFiscal
                                    worksheet.Cells("T" & i).Value = ValorFiscal
                                    worksheet.Cells("U" & i).Value = PesoFiscal_Dif
                                    worksheet.Cells("V" & i).Value = ValorFiscal_Dif

                                    worksheet.Cells("X" & i).Value = PesoFinanceiro
                                    worksheet.Cells("Y" & i).Value = ValorFinanceiro
                                    worksheet.Cells("Z" & i).Value = PesoFinanceiro_Dif
                                    worksheet.Cells("AA" & i).Value = ValorFinanceiro_Dif

                                    worksheet.Cells(String.Format("I{0}:K{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("L{0}:M{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("N{0}:O{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("P{0}:Q{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("S{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("T{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("U{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("V{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("X{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("Y{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells(String.Format("Z{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells(String.Format("AA{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                    worksheet.Cells("I" & i & ":AA" & i).Style.Font.Bold = True

                                    Contratado = 0
                                    Entregue = 0
                                    Pago = 0
                                    APagar = 0
                                    ValorTotal = 0
                                    ValorPago = 0

                                    PesoFiscal = 0
                                    ValorFiscal = 0
                                    PesoFiscal_Dif = 0
                                    ValorFiscal_Dif = 0
                                    PesoFinanceiro = 0
                                    ValorFinanceiro = 0
                                    PesoFinanceiro_Dif = 0
                                    ValorFinanceiro_Dif = 0

                                    Using range = worksheet.Cells("I" & i & ":I" & i)
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Left.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                                    End Using

                                    Using range = worksheet.Cells("J" & i & ":AA" & i)
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                                    End Using

                                    Using range = worksheet.Cells("AB" & i & ":AB" & i)
                                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Right.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                                    End Using

                                    i = i + 1
                                End If

                                worksheet.Cells("A" & i & ":H" & i).Merge = True
                                worksheet.Cells("A" & i).Value = "      " & dr("Produto") & "-" & dr("NomeDoProduto")
                                worksheet.Cells("A" & i & ":H" & i).Style.Font.Bold = True
                                QuebraProduto = dr("Produto")
                                i = i + 1
                            End If

                            worksheet.Cells("B" & i).Value = dr("NomeCliente")
                            worksheet.Cells("C" & i).Value = dr("CidadeCliente")
                            worksheet.Cells("D" & i).Value = dr("UFCliente")
                            worksheet.Cells("E" & i).Value = dr("Pedido")
                            worksheet.Cells("F" & i).Value = Format(CDate(dr("DataPedido")), "dd-MM-yyyy")

                            worksheet.Cells("G" & i).Value = dr("OPeracao")
                            worksheet.Cells("H" & i).Value = dr("SubOperacao")
                            worksheet.Cells("I" & i).Value = dr("Contratado")

                            worksheet.Cells("J" & i).Value = dr("Contratado") - dr("AEntregar")
                            worksheet.Cells("K" & i).Value = dr("AEntregar")

                            worksheet.Cells("L" & i).Value = dr("UnitarioOficial")
                            worksheet.Cells("M" & i).Value = dr("TotalOficial")

                            worksheet.Cells("N" & i).Value = dr("Pago")
                            worksheet.Cells("O" & i).Value = dr("Contratado") - dr("Pago")

                            worksheet.Cells("P" & i).Value = dr("Baixado_SaldoOficial")
                            worksheet.Cells("Q" & i).Value = dr("TotalLiquidoOficial") - dr("Baixado_SaldoOficial")

                            'worksheet.Cells("Q" & i).Value = dr("TotalOficial") - dr("Baixado_SaldoOficial")

                            If dr("Complementacoes") = 0 Then
                                worksheet.Cells("S" & i).Value = dr("PesoFiscal")
                                worksheet.Cells("T" & i).Value = dr("ValorLiquidoDeNota")
                            Else
                                worksheet.Cells("S" & i).Value = dr("Complementacoes")
                                worksheet.Cells("T" & i).Value = dr("TotalOficial")
                            End If

                            worksheet.Cells("U" & i).Value = (dr("Contratado") - dr("AEntregar")) - dr("PesoFiscal")
                            If worksheet.Cells("S" & i).Value <> 0 And worksheet.Cells("T" & i).Value <> 0 Then
                                worksheet.Cells("V" & i).Value = (worksheet.Cells("T" & i).Value / worksheet.Cells("S" & i).Value) * worksheet.Cells("U" & i).Value
                            End If

                            worksheet.Cells("X" & i).Value = dr("Contratado") - dr("AEntregar")
                            If worksheet.Cells("X" & i).Value <> 0 And worksheet.Cells("S" & i).Value <> 0 Then
                                worksheet.Cells("Y" & i).Value = (worksheet.Cells("T" & i).Value / worksheet.Cells("S" & i).Value) * worksheet.Cells("X" & i).Value
                            End If

                            worksheet.Cells("Z" & i).Value = (dr("Contratado") - dr("AEntregar")) - dr("Pago")
                            If worksheet.Cells("S" & i).Value <> 0 And worksheet.Cells("T" & i).Value <> 0 Then
                                worksheet.Cells("AA" & i).Value = (worksheet.Cells("T" & i).Value / worksheet.Cells("S" & i).Value) * worksheet.Cells("Z" & i).Value
                            End If

                            worksheet.Cells("AB" & i).Value = dr("Adiantamento")

                            '----Zebrado------------------------------------------

                            If i Mod 2 = 0 Then
                                Using range = worksheet.Cells("A" & i & ":M" & i)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                                Using range = worksheet.Cells("N" & i & ":Q" & i)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 153))
                                End Using
                                Using range = worksheet.Cells("S" & i & ":V" & i)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 204))
                                End Using
                                Using range = worksheet.Cells("X" & i & ":AB" & i)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 255))
                                End Using
                            End If

                            worksheet.Cells(String.Format("I{0}:K{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells(String.Format("L{0}:M{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("N{0}:O{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells(String.Format("P{0}:Q{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("S{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells(String.Format("T{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("U{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells(String.Format("V{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("X{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells(String.Format("Y{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("Z{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells(String.Format("AA{0}:AB{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            '---Totalizadores--------------------------------------

                            Contratado = Contratado + dr("Contratado")
                            Entregue = Entregue + (dr("Contratado") - dr("AEntregar"))

                            ValorTotal = ValorTotal + dr("TotalOficial")
                            ValorPago = ValorPago + dr("Baixado_SaldoOficial")

                            Pago = Pago + dr("Pago")
                            APagar = APagar + (dr("Contratado") - dr("Pago"))

                            PesoFiscal = PesoFiscal + worksheet.Cells("S" & i).Value
                            ValorFiscal = ValorFiscal + worksheet.Cells("T" & i).Value
                            PesoFiscal_Dif = PesoFiscal_Dif + worksheet.Cells("U" & i).Value
                            ValorFiscal_Dif = ValorFiscal_Dif + worksheet.Cells("V" & i).Value

                            PesoFinanceiro = PesoFinanceiro + worksheet.Cells("X" & i).Value
                            ValorFinanceiro = ValorFinanceiro + worksheet.Cells("Y" & i).Value
                            PesoFinanceiro_Dif = PesoFinanceiro_Dif + worksheet.Cells("Z" & i).Value
                            ValorFinanceiro_Dif = ValorFinanceiro_Dif + worksheet.Cells("AA" & i).Value

                            i = i + 1
                        Next

                        'worksheet.Cells(String.Format("I{0}", i)).Formula = String.Format("=SUM(I12:I{0})", i - 1)

                        'worksheet.Cells("I" & i).Value = "******5*******"
                        worksheet.Cells("I" & i).Value = Contratado
                        worksheet.Cells("J" & i).Value = Entregue
                        worksheet.Cells("K" & i).Value = Contratado - Entregue

                        worksheet.Cells("M" & i).Value = ValorTotal
                        worksheet.Cells("N" & i).Value = Pago
                        worksheet.Cells("O" & i).Value = APagar

                        worksheet.Cells("P" & i).Value = ValorPago
                        worksheet.Cells("Q" & i).Value = ValorTotal - ValorPago

                        worksheet.Cells("S" & i).Value = PesoFiscal
                        worksheet.Cells("T" & i).Value = ValorFiscal
                        worksheet.Cells("U" & i).Value = PesoFiscal_Dif
                        worksheet.Cells("V" & i).Value = ValorFiscal_Dif

                        worksheet.Cells("X" & i).Value = PesoFinanceiro
                        worksheet.Cells("Y" & i).Value = ValorFinanceiro
                        worksheet.Cells("Z" & i).Value = PesoFinanceiro_Dif
                        worksheet.Cells("AA" & i).Value = ValorFinanceiro_Dif

                        worksheet.Cells(String.Format("I{0}:K{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                        worksheet.Cells(String.Format("L{0}:M{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("N{0}:O{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                        worksheet.Cells(String.Format("P{0}:Q{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("S{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                        worksheet.Cells(String.Format("T{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("U{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                        worksheet.Cells(String.Format("V{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("X{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                        worksheet.Cells(String.Format("Y{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("Z{0}", i)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                        worksheet.Cells(String.Format("AA{0}:AB{0}", i)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        worksheet.Cells("I" & i & ":AA" & i).Style.Font.Bold = True

                        Using range = worksheet.Cells("I" & i & ":I" & i)
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                            range.Style.Border.Left.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Left.Color.SetColor(Color.FromArgb(0, 0, 0))
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                        End Using

                        Using range = worksheet.Cells("J" & i & ":AA" & i)
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                        End Using

                        Using range = worksheet.Cells("AB" & i & ":AB" & i)
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Top.Color.SetColor(Color.FromArgb(0, 0, 0))
                            range.Style.Border.Right.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Right.Color.SetColor(Color.FromArgb(0, 0, 0))
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                        End Using

                        Contratado = 0
                        Entregue = 0
                        Pago = 0
                        APagar = 0
                        ValorTotal = 0
                        ValorPago = 0

                        PesoFiscal = 0
                        ValorFiscal = 0
                        PesoFiscal_Dif = 0
                        ValorFiscal_Dif = 0

                        PesoFinanceiro = 0
                        ValorFinanceiro = 0
                        PesoFinanceiro_Dif = 0
                        ValorFinanceiro_Dif = 0

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linha
                        worksheet.Cells("A8:AB8").Style.Font.Bold = True
                        worksheet.View.FreezePanes(9, 1)
                        '.Range("A9").Select()
                        ''.ActiveWindow.FreezePanes = True
                        package.Save()
                    End Using
                End Using

                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSetExcel() As DataSet
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy-MM-dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy-MM-dd")

        Dim ds As DataSet
        AtualizaPosicaoDePedidos()

        If cbVencimento.Checked Then
            Sql &= "SELECT Emp.Reduzido AS Empresa," & vbCrLf &
                   "      Emp.Cliente_Id AS CNPJEmpresa," & vbCrLf &
                   "      Emp.Endereco_Id AS EndEmpresa," & vbCrLf &
                   "      Emp.Nome AS NomeEmpresa, " & vbCrLf &
                   "	   Emp.Cidade AS CidadeEmpresa," & vbCrLf &
                   "	   Emp.Estado AS UFEmpresa, " & vbCrLf &
                   "	   cli.Cliente_Id AS Cliente," & vbCrLf &
                   "	   cli.Endereco_Id AS EndCliente, " & vbCrLf &
                   "	   cli.Nome AS NomeCliente," & vbCrLf &
                   "	   cli.Cidade AS CidadeCliente, " & vbCrLf &
                   "	   cli.Estado AS UFCliente, " & vbCrLf &
                   "	   sOP.Operacao_Id AS Operacao, " & vbCrLf &
                   "	   sOP.SubOperacoes_Id AS SubOperacao, " & vbCrLf &
                   "	   sOP.Descricao AS NomeOperacao, " & vbCrLf &
                   "	   p.Moeda, " & vbCrLf &
                   "	   m.Descricao AS NomeMoeda, " & vbCrLf &
                   "	   p.Indexador, " & vbCrLf &
                   "	   ind.Descricao AS NomeIndexador, " & vbCrLf &
                   "	   p.Finalidade, " & vbCrLf &
                   "	   fin.Descricao AS NomeFinalidade, " & vbCrLf &
                   "	   p.FreteCIFFOB, " & vbCrLf &
                   "	   p.CondicaoPagamento, " & vbCrLf &
                   "	   pgto.Descricao AS NomeCondicaoDePagamento, " & vbCrLf &
                   "	   p.DataPedido, " & vbCrLf &
                   "	   p.DataEntrega, " & vbCrLf &
                   "	   #Posicao.Pedido_Id AS Pedido, " & vbCrLf &
                   "	   #Posicao.Produto, " & vbCrLf &
                   "	   prd.Nome AS NomeDoProduto, " & vbCrLf &
                   "	   #Posicao.Contratado, " & vbCrLf &
                   "	   #Posicao.UnitarioOficial, " & vbCrLf &
                   "	   #Posicao.TotalOficial, " & vbCrLf &
                   "	   #Posicao.UnitarioMoeda, " & vbCrLf &
                   "	   #Posicao.TotalMoeda, " & vbCrLf &
                   "	   #Posicao.Laudo As Romaneio, " & vbCrLf &
                   "	   #Posicao.Entregue AS PesoFiscal, " & vbCrLf &
                   "	   #Posicao.AEntregar, " & vbCrLf &
                   "	   #Posicao.QuantidadeFixado AS Fixado, " & vbCrLf &
                   "	   #Posicao.AFixar, " & vbCrLf &
                   "	   isnull(#Posicao.Pago, 0) as Pago, " & vbCrLf &
                   "	   #Posicao.PagoNaoRecebido AS PagoNaoEntregue, " & vbCrLf &
                   "	   #Posicao.RecebidoNaoPago AS EntregueNaoPago, " & vbCrLf &
                   "	   #Posicao.Adiantamento, " & vbCrLf &
                   "	   p.UsuarioInclusao, " & vbCrLf &
                   "	   p.UsuarioAlteracao, " & vbCrLf &
                   "	   p.UsuarioLiberacao, " & vbCrLf &
                   "	   #Posicao.Programado_PagarOficial, " & vbCrLf &
                   "	   #Posicao.Baixado_PagarOficial, " & vbCrLf &
                   "	   #Posicao.Programado_PagarMoeda, " & vbCrLf &
                   "	   #Posicao.Baixado_PagarMoeda, " & vbCrLf &
                   "	   #Posicao.Programado_ReceberOficial, " & vbCrLf &
                   "	   #Posicao.Programado_ReceberMoeda, " & vbCrLf &
                   "	   #Posicao.Baixado_ReceberOficial, " & vbCrLf &
                   "	   #Posicao.Baixado_ReceberMoeda, " & vbCrLf &
                   "	   #Posicao.Programado_SaldoOficial, " & vbCrLf &
                   "	   #Posicao.Baixado_SaldoOficial, " & vbCrLf &
                   "	   #Posicao.Programado_SaldoMoeda, " & vbCrLf &
                   "	   #Posicao.Baixado_SaldoMoeda, " & vbCrLf &
                   "	   #Posicao.OrigemDestino, " & vbCrLf &
                   "	   #Posicao.EntradaSaida, " & vbCrLf &
                   "	   #Posicao.MoedaPedido, " & vbCrLf &
                   "	   #Posicao.ValorFixadoOficial , " & vbCrLf &
                   "	   #Posicao.ValorFixadoMoeda, " & vbCrLf &
                   "	   #Posicao.ValorLiquidoDeNota, " & vbCrLf &
                   "      #Posicao.Complementacoes,    " & vbCrLf &
                   "  isnull(#Posicao.TotalLiquidoOficial, 0) as TotalLiquidoOficial " & vbCrLf &
                   "  FROM #Posicao" & vbCrLf &
                   " INNER JOIN Pedidos p" & vbCrLf &
                   "    ON p.Empresa_Id     = #Posicao.Empresa_Id" & vbCrLf &
                   "   AND p.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
                   "   AND p.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf

            If FinanceiroNovo Then
                Sql &= " INNER JOIN Titulos cP " & vbCrLf &
                    "    ON cP.Empresa		= #Posicao.Empresa_Id " & vbCrLf &
                       "   AND cP.EndEmpresa	= #Posicao.EndEmpresa_Id " & vbCrLf &
                       "   AND cP.Pedido        = #Posicao.Pedido_Id " & vbCrLf
            Else
                Sql &= " INNER JOIN ContasAPagar cP" & vbCrLf &
                       "	ON cP.EmpresaPedido     = #Posicao.Empresa_Id" & vbCrLf &
                       "   AND cP.EndEmpresaPedido = #Posicao.EndEmpresa_Id" & vbCrLf &
                       "   AND cP.Pedido           = #Posicao.Pedido_Id" & vbCrLf
            End If

            Sql &= " INNER JOIN Clientes cli" & vbCrLf &
                   "	ON cli.Cliente_Id   = p.Cliente" & vbCrLf &
                   "   AND cli.Endereco_Id = p.EndCliente" & vbCrLf &
                   "  LEFT JOIN Produtos prd" & vbCrLf &
                   "	ON prd.Produto_Id = #Posicao.Produto " & vbCrLf &
                   "  LEFT JOIN Pagamentos pgto" & vbCrLf &
                   "	ON pgto.Pagamento_Id = p.CondicaoPagamento" & vbCrLf &
                   "  LEFT JOIN Indexadores ind" & vbCrLf &
                   "	ON ind.Indexador_Id = p.Indexador " & vbCrLf &
                   "  LEFT JOIN Moedas m" & vbCrLf &
                   "	ON m.Moeda_Id = p.Moeda" & vbCrLf &
                   "  LEFT JOIN Finalidades fin" & vbCrLf &
                   "	ON fin.Finalidade_Id = p.Finalidade" & vbCrLf &
                   "  LEFT JOIN SubOperacoes sOp" & vbCrLf &
                   "	ON sOp.Operacao_Id      = p.Operacao " & vbCrLf &
                   "   AND sOp.SubOperacoes_Id = p.SubOperacao " & vbCrLf &
                   "  LEFT JOIN Clientes AS Emp " & vbCrLf &
                   "	ON Emp.Cliente_Id = p.Empresa_Id" & vbCrLf &
                   "   AND Emp.Endereco_Id = p.EndEmpresa_Id " & vbCrLf &
                   "  LEFT JOIN ComprasXProdutos cXp" & vbCrLf &
                   "	ON cXp.Produto_Id = p.Carteira" & vbCrLf

            If FinanceiroNovo Then
                Sql &= " WHERE (cP.Reprogramacao between '2013-12-02' AND '2013-12-12') " & vbCrLf
            Else
                Sql &= " WHERE (cP.Prorrogacao between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf
            End If

            Sql &= "   AND (sOP.Financeiro = 'S' or sOP.classe = '" & eClassesOperacoes.AFIXAR.ToString & "') " & vbCrLf &
                "   And left(#Posicao.Produto, 5)  in ('10101', '10102', '10201', '10202', '10203') " & vbCrLf

            If rdAbertos.Checked Then
                Sql &= "   And ((#Posicao.Pago <> (#Posicao.Contratado - #Posicao.AEntregar)) or (#Posicao.AEntregar <> 0))  " & vbCrLf
            End If
        Else
            Sql &= "SELECT Emp.Reduzido AS Empresa," & vbCrLf &
                   "       Emp.Cliente_Id AS CNPJEmpresa," & vbCrLf &
                   "       Emp.Endereco_Id AS EndEmpresa," & vbCrLf &
                   "       Emp.Nome AS NomeEmpresa," & vbCrLf &
                   "	   Emp.Cidade AS CidadeEmpresa," & vbCrLf &
                   "	   Emp.Estado AS UFEmpresa," & vbCrLf &
                   "	   cli.Cliente_Id AS Cliente," & vbCrLf &
                   "	   cli.Endereco_Id AS EndCliente," & vbCrLf &
                   "	   cli.Nome AS NomeCliente," & vbCrLf &
                   "	   cli.Cidade AS CidadeCliente," & vbCrLf &
                   "	   cli.Estado AS UFCliente," & vbCrLf &
                   "	   sOp.Operacao_Id AS Operacao," & vbCrLf &
                   "	   sOp.SubOperacoes_Id AS SubOperacao," & vbCrLf &
                   "	   sOp.Descricao AS NomeOperacao," & vbCrLf &
                   "	   p.Moeda," & vbCrLf &
                   "	   m.Descricao AS NomeMoeda," & vbCrLf &
                   "	   p.Indexador," & vbCrLf &
                   "	   ind.Descricao AS NomeIndexador," & vbCrLf &
                   "	   p.Finalidade," & vbCrLf &
                   "	   fin.Descricao AS NomeFinalidade," & vbCrLf &
                   "	   p.FreteCIFFOB," & vbCrLf &
                   "	   p.CondicaoPagamento," & vbCrLf &
                   "	   pgto.Descricao AS NomeCondicaoDePagamento," & vbCrLf &
                   "	   p.DataPedido," & vbCrLf &
                   "	   p.DataEntrega," & vbCrLf &
                   "	   #Posicao.Pedido_Id AS Pedido," & vbCrLf &
                   "	   #Posicao.Produto," & vbCrLf &
                   "	   prd.Nome AS NomeDoProduto," & vbCrLf &
                   "	   #Posicao.Contratado," & vbCrLf &
                   "	   #Posicao.UnitarioOficial," & vbCrLf &
                   "	   isnull(#Posicao.TotalOficial, 0) as TotalOficial," & vbCrLf &
                   "	   #Posicao.UnitarioMoeda," & vbCrLf &
                   "	   #Posicao.TotalMoeda," & vbCrLf &
                   "	   #Posicao.Laudo As Romaneio," & vbCrLf &
                   "	   isnull(#Posicao.Entregue, 0) AS PesoFiscal," & vbCrLf &
                   "	   isnull(#Posicao.AEntregar, 0) as AEntregar," & vbCrLf &
                   "	   isnull(#Posicao.QuantidadeFixado, 0) AS Fixado," & vbCrLf &
                   "	   isnull(#Posicao.AFixar, 0) as AFixar," & vbCrLf &
                   "	   isnull(#Posicao.Pago, 0) as Pago," & vbCrLf &
                   "	   isnull(#Posicao.PagoNaoRecebido, 0) AS PagoNaoEntregue," & vbCrLf &
                   "	   isnull(#Posicao.RecebidoNaoPago, 0) AS EntregueNaoPago," & vbCrLf &
                   "	   isnull(#Posicao.Adiantamento, 0) as Adiantamento," & vbCrLf &
                   "	   p.UsuarioInclusao," & vbCrLf &
                   "	   p.UsuarioAlteracao," & vbCrLf &
                   "	   p.UsuarioLiberacao," & vbCrLf &
                   "	   #Posicao.Programado_PagarOficial," & vbCrLf &
                   "	   #Posicao.Baixado_PagarOficial," & vbCrLf &
                   "	   #Posicao.Programado_PagarMoeda," & vbCrLf &
                   "	   #Posicao.Baixado_PagarMoeda," & vbCrLf &
                   "	   #Posicao.Programado_ReceberOficial," & vbCrLf &
                   "	   #Posicao.Programado_ReceberMoeda," & vbCrLf &
                   "	   #Posicao.Baixado_ReceberOficial," & vbCrLf &
                   "	   #Posicao.Baixado_ReceberMoeda," & vbCrLf &
                   "	   #Posicao.Programado_SaldoOficial," & vbCrLf &
                   "	   #Posicao.Baixado_SaldoOficial," & vbCrLf &
                   "	   #Posicao.Programado_SaldoMoeda," & vbCrLf &
                   "	   #Posicao.Baixado_SaldoMoeda," & vbCrLf &
                   "	   #Posicao.OrigemDestino," & vbCrLf &
                   "	   #Posicao.EntradaSaida," & vbCrLf &
                   "	   #Posicao.MoedaPedido," & vbCrLf &
                   "	   #Posicao.ValorFixadoOficial," & vbCrLf &
                   "	   #Posicao.ValorFixadoMoeda," & vbCrLf &
                   "	   IsNull(#Posicao.ValorLiquidoDeNota,0) as ValorLiquidoDeNota," & vbCrLf &
                   "	   isnull(#Posicao.Complementacoes, 0) as Complementacoes," & vbCrLf &
                   "       isnull(#Posicao.TotalLiquidoOficial, 0) as TotalLiquidoOficial" & vbCrLf &
                   "  FROM #Posicao" & vbCrLf &
                   " INNER JOIN Pedidos p" & vbCrLf &
                   "	ON p.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
                   "   AND p.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
                   "   AND p.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf &
                   " INNER JOIN Clientes cli" & vbCrLf &
                   "	ON cli.Cliente_Id  = p.Cliente" & vbCrLf &
                   "   AND cli.Endereco_Id = p.EndCliente" & vbCrLf &
                   "  LEFT JOIN Indexadores ind" & vbCrLf &
                   "	ON ind.Indexador_Id = p.Indexador" & vbCrLf &
                   "  LEFT JOIN Pagamentos pgto" & vbCrLf &
                   "	ON pgto.Pagamento_Id = p.CondicaoPagamento" & vbCrLf &
                   "  LEFT JOIN Moedas m" & vbCrLf &
                   "	ON m.Moeda_Id = p.Moeda" & vbCrLf &
                   "  LEFT JOIN Finalidades fin" & vbCrLf &
                   "	ON fin.Finalidade_Id = p.Finalidade" & vbCrLf &
                   "  LEFT JOIN SubOperacoes sOp" & vbCrLf &
                   "	ON sOp.Operacao_Id     = p.Operacao" & vbCrLf &
                   "   AND sOp.SubOperacoes_Id = p.SubOperacao" & vbCrLf &
                   "  LEFT JOIN Clientes AS Emp" & vbCrLf &
                   "	ON Emp.Cliente_Id  = p.Empresa_Id" & vbCrLf &
                   "   AND Emp.Endereco_Id = p.EndEmpresa_Id" & vbCrLf &
                   "  LEFT JOIN ComprasXProdutos cXp" & vbCrLf &
                   "	ON cXp.Produto_Id = p.Carteira" & vbCrLf &
                   "  LEFT JOIN Produtos prd" & vbCrLf &
                   "    ON prd.Produto_Id = #Posicao.Produto" & vbCrLf &
                   " WHERE (p.DataPedido between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "   AND (sOp.Financeiro = 'S' or sOp.classe = '" & eClassesOperacoes.AFIXAR.ToString & "')" & vbCrLf &
                   "   And left(#Posicao.Produto, 5)  in ('10101', '10102', '10201', '10202', '10203', '10220')" & vbCrLf

            Sql &= "        --AND P.Pedido_Id     = 7332 " & vbCrLf

            If rdAbertos.Checked Then
                Sql &= "   And ((#Posicao.Pago <> (#Posicao.Contratado - #Posicao.AEntregar)) or (#Posicao.AEntregar <> 0))" & vbCrLf
            End If

            Sql &= " Order By Emp.Reduzido,  #Posicao.EntradaSaida, #Posicao.Produto, cli.Nome, #Posicao.Pedido_Id" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(Sql, "Pedidos")
        Return ds
    End Function

    Private Sub AtualizaPosicaoDePedidos()

        '************PARTE 1 - LINHA 775****************************
        Sql = "SELECT p.Empresa_Id," & vbCrLf &
              "       p.EndEmpresa_Id," & vbCrLf &
              "       p.Pedido_Id," & vbCrLf &
              "       p.OrigemDestino," & vbCrLf &
              "       p.Moeda AS MoedaPedido," & vbCrLf &
              "       sOp.EntradaSaida," & vbCrLf &
              "       pXi.Produto_Id           AS Produto," & vbCrLf &
              "       convert(Decimal(18,4),0) AS Contratado," & vbCrLf &
              "       convert(Decimal(18,4),0) AS Laudo," & vbCrLf &
              "       convert(Decimal(18,4),0) AS Entregue," & vbCrLf &
              "	      convert(Decimal(18,4),0) AS AEntregar," & vbCrLf &
              "	      convert(Decimal(18,4),0) AS QuantidadeFixado," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS ValorFixadoOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS ValorFixadoMoeda," & vbCrLf &
              "	      convert(Decimal(18,4),0) AS AFixar," & vbCrLf &
              "	      convert(Decimal(18,4),0) AS Pago," & vbCrLf &
              "	      convert(Decimal(18,4),0) AS PagoNaoRecebido," & vbCrLf &
              "	      convert(Decimal(18,4),0) AS RecebidoNaoPago," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Adiantamento," & vbCrLf &
              "	      convert(Decimal(18,10),0)AS UnitarioOficial," & vbCrLf &
              "	      convert(Decimal(18,10),0)AS UnitarioMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS TotalOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS TotalMoeda," & vbCrLf &
              "       convert(Decimal(18,2),0) AS TotalLiquidoOficial," & vbCrLf &
              "       convert(Decimal(18,2),0) AS TotalLiquidoMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Programado_PagarOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Baixado_PagarOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Programado_PagarMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Baixado_PagarMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Programado_ReceberOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Baixado_ReceberOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Programado_ReceberMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Baixado_ReceberMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Programado_SaldoOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Baixado_SaldoOficial," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Programado_SaldoMoeda," & vbCrLf &
              "	      convert(Decimal(18,2),0) AS Baixado_SaldoMoeda," & vbCrLf &
              "       convert(Decimal(18,2),0) AS ValorLiquidoDeNota," & vbCrLf &
              "       convert(Decimal(18,2),0) AS Complementacoes" & vbCrLf &
              "  into #Posicao" & vbCrLf &
              "  FROM Pedidos p" & vbCrLf &
              " INNER JOIN PedidoXItem pXi" & vbCrLf &
              "    ON pXi.Empresa_Id       = p.Empresa_Id" & vbCrLf &
              "   AND pXi.EndEmpresa_Id    = p.EndEmpresa_Id" & vbCrLf &
              "   AND pXi.Pedido_Id        = p.Pedido_Id" & vbCrLf &
              "  INNER JOIN Produtos Prd" & vbCrLf &
              "    ON Prd.Produto_Id = pXi.Produto_Id" & vbCrLf &
              " INNER JOIN SubOperacoes sOp" & vbCrLf &
              "    ON sOp.Operacao_Id     = p.Operacao" & vbCrLf &
              "   AND sOp.SubOperacoes_Id = p.SubOperacao" & vbCrLf &
              " WHERE p.Empresa_Id    ='" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
              "   AND p.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
              "   And Prd.Agrupar     ='N'" & vbCrLf &
              "   And p.Situacao      = 1 " & vbCrLf &
              "   And p.DataPedido between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf

        'Atualiza Quantidade Contratada Itens do Pedido----
        'PARTE 2 - LINHA 798

        Sql &= "Update #Posicao SET" & vbCrLf &
               "   #Posicao.Contratado      = consulta.Quantidade," & vbCrLf &
               "   #Posicao.TotalOficial    = consulta.TotalOficial," & vbCrLf &
               "   #Posicao.TotalMoeda      = consulta.TotalMoeda," & vbCrLf &
               "   #Posicao.TotalLiquidoOficial = consulta.LiqOficial," & vbCrLf &
               "   #Posicao.TotalLiquidoMoeda   = consulta.LiqMoeda" & vbCrLf &
               "   FROM (SELECT pXi.Empresa_Id," & vbCrLf &
               "                pXi.EndEmpresa_Id," & vbCrLf &
               "                pXi.Pedido_Id," & vbCrLf &
               "                pXi.Produto_Id," & vbCrLf &
               "                encLiq.ValorOficial as LiqOficial," & vbCrLf &
               "                encLiq.ValorMoeda   as LiqMoeda,  " & vbCrLf &
               "			    SUM(case" & vbCrLf &
               "			          when sOp.Classe <> '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf &
               "			            THEN CASE" & vbCrLf &
               "			                   WHEN pXi.TipoDeLancamento = 'E'" & vbCrLf &
               "			                     THEN pXi.TotalOficial * - 1" & vbCrLf &
               "			                     ELSE pXi.TotalOficial" & vbCrLf &
               "                             End" & vbCrLf &
               "			            Else 0" & vbCrLf &
               "			        End)  AS TotalOficial," & vbCrLf &
               "			    SUM(case" & vbCrLf &
               "			          when sOp.Classe <> '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf &
               "			            THEN CASE" & vbCrLf &
               "			                   WHEN pXi.TipoDeLancamento = 'E'" & vbCrLf &
               "			                     THEN pXi.TotalMoeda * - 1" & vbCrLf &
               "			                     ELSE pXi.TotalMoeda" & vbCrLf &
               "                             End" & vbCrLf &
               "			            Else 0" & vbCrLf &
               "			        End) AS TotalMoeda," & vbCrLf &
               "			    SUM(CASE" & vbCrLf &
               "			          WHEN pXi.TipoDeLancamento = 'E'" & vbCrLf &
               "			            THEN pXi.Quantidade * - 1" & vbCrLf &
               "			            ELSE pXi.Quantidade" & vbCrLf &
               "			        END) AS Quantidade" & vbCrLf &
               "			FROM PedidoXItemXLancamento pXi" & vbCrLf &
               "		   INNER JOIN Pedidos p" & vbCrLf &
               "		      ON p.Empresa_Id    = pXi.Empresa_Id" & vbCrLf &
               "			 AND p.EndEmpresa_Id = pXi.EndEmpresa_Id" & vbCrLf &
               "			 AND p.Pedido_Id     = pXi.Pedido_Id" & vbCrLf &
               "		   INNER JOIN SubOperacoes sOp" & vbCrLf &
               "		      ON sOp.Operacao_Id     = p.Operacao" & vbCrLf &
               "			 AND sOp.SubOperacoes_Id = p.SubOperacao" & vbCrLf &
               "           Inner Join(Select Empresa_id," & vbCrLf &
               "                             EndEmpresa_id," & vbCrLf &
               "                             Pedido_Id," & vbCrLf &
               "                             Produto_id," & vbCrLf &
               "                             ValorOficial," & vbCrLf &
               "                             ValorMoeda" & vbCrLf &
               "                        from Pedidosxencargos" & vbCrLf &
               "                       where Encargo_id = 'LIQUIDO'" & vbCrLf &
               "                      ) As EncLiq" & vbCrLf &
               "              ON pXi.Empresa_Id    = EncLiq.Empresa_Id    " & vbCrLf &
               "	         AND pXi.EndEmpresa_Id = EncLiq.EndEmpresa_Id  " & vbCrLf &
               "	         AND pXi.Pedido_Id     = EncLiq.Pedido_Id      " & vbCrLf &
               "	         AND pXi.Produto_id    = EncLiq.Produto_id     " & vbCrLf &
               "		   WHERE pXi.Empresa_Id    ='" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
               "		     AND pXi.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
               "		     And p.Situacao        = 1" & vbCrLf &
               "		   GROUP BY pXi.Empresa_Id, pXi.EndEmpresa_Id, pXi.Pedido_Id, pXi.Produto_Id, encLiq.ValorOficial, encLiq.ValorMoeda " & vbCrLf &
               "		 ) AS consulta" & vbCrLf &
               "  INNER JOIN #Posicao" & vbCrLf &
               "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
               "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	AND	consulta.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf &
               "	AND	consulta.Produto_Id    = #Posicao.Produto" & vbCrLf


        'Atualiza Complementações de Preco ---------
        'PARTE 3 - LINHA 833
        Sql &= "Update #Posicao SET" & vbCrLf &
               "	#Posicao.Complementacoes = consulta.Quantidade," & vbCrLf &
               "	#Posicao.TotalOficial    = consulta.TotalOficial," & vbCrLf &
               "	#Posicao.TotalMoeda      = consulta.TotalMoeda" & vbCrLf &
               "  FROM (SELECT pXi.Empresa_Id," & vbCrLf &
               "               pXi.EndEmpresa_Id," & vbCrLf &
               "               pXi.Pedido_Id," & vbCrLf &
               "               pXi.Produto_Id," & vbCrLf &
               "			   SUM(CASE" & vbCrLf &
               "			         WHEN sOp.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
               "			           THEN pXiXf.TotalOficial" & vbCrLf &
               "			           ELSE 0.00" & vbCrLf &
               "			       END) AS TotalOficial," & vbCrLf &
               "			   SUM(CASE" & vbCrLf &
               "			         WHEN sOp.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
               "			           THEN pXiXf.TotalMoeda" & vbCrLf &
               "			           ELSE 0.00" & vbCrLf &
               "			       END) AS TotalMoeda," & vbCrLf &
               "			   SUM(CASE" & vbCrLf &
               "			         WHEN sOp.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
               "			           THEN pXiXf.Quantidade" & vbCrLf &
               "			           ELSE 0.00" & vbCrLf &
               "			       END) AS Quantidade" & vbCrLf &
               "		  FROM Pedidos p" & vbCrLf &
               "	     INNER JOIN PedidoXItem pXi" & vbCrLf &
               "			ON pXi.Empresa_Id    = p.Empresa_Id" & vbCrLf &
               "		   AND pXi.EndEmpresa_Id = p.EndEmpresa_Id" & vbCrLf &
               "		   AND pXi.Pedido_Id     = p.Pedido_Id" & vbCrLf &
               "	     INNER JOIN VW_PedidosXItensXFixacoes pXiXf" & vbCrLf &
               "			ON pXiXf.Empresa_Id    = pXi.Empresa_Id" & vbCrLf &
               "		   AND pXiXf.EndEmpresa_Id = pXi.EndEmpresa_Id" & vbCrLf &
               "		   AND pXiXf.Pedido_Id     = pXi.Pedido_Id" & vbCrLf &
               "		   AND pXiXf.Produto_Id    = pXi.Produto_Id" & vbCrLf &
               "	     INNER JOIN SubOperacoes sOp" & vbCrLf &
               "			ON sOp.Operacao_Id     = pXiXf.Operacao" & vbCrLf &
               "		   AND sOp.SubOperacoes_Id = pXiXf.SubOperacao" & vbCrLf &
               "		 WHERE p.Empresa_Id    ='" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
               "		   AND p.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
               "		   And p.Situacao      = 1" & vbCrLf &
               "		 GROUP BY pXi.Empresa_Id, pXi.EndEmpresa_Id, pXi.Pedido_Id, pXi.Produto_Id" & vbCrLf &
               "		) AS consulta" & vbCrLf &
               "	INNER JOIN #Posicao" & vbCrLf &
               "	   ON consulta.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
               "	  AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	  AND consulta.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf &
               "	  AND consulta.Produto_Id    = #Posicao.Produto" & vbCrLf &
               "	INNER JOIN Pedidos AS ConsultasXPedidos" & vbCrLf &
               "	   ON ConsultasXPedidos.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
               "	  AND ConsultasXPedidos.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	  AND ConsultasXPedidos.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf &
               "	INNER JOIN SubOperacoes AS ConsultasXSubOperacoes" & vbCrLf &
               "	   ON ConsultasXSubOperacoes.Operacao_Id     = ConsultasXPedidos.Operacao" & vbCrLf &
               "	  AND ConsultasXSubOperacoes.SubOperacoes_Id = ConsultasXPedidos.SubOperacao" & vbCrLf &
               "	WHERE (ConsultasXSubOperacoes.Classe = '" & eClassesOperacoes.AFIXAR.ToString & "')" & vbCrLf


        'Atualiza Quantidade Fixada Pedidos X Itens X Fixacoes ----
        'PARTE 4 - LINHA 867
        Sql &= "Update #Posicao SET" & vbCrLf &
               "	#Posicao.QuantidadeFixado   = consulta.Quantidade," & vbCrLf &
               "	#Posicao.ValorFixadoOficial = consulta.TotalOficial," & vbCrLf &
               "	#Posicao.ValorFixadoMoeda   = consulta.TotalMoeda" & vbCrLf &
               "  FROM (SELECT pXiXf.Empresa_Id," & vbCrLf &
               "	           pXiXf.EndEmpresa_Id," & vbCrLf &
               "	           pXiXf.Pedido_Id," & vbCrLf &
               "			   SUM(pXiXf.Quantidade) AS Quantidade," & vbCrLf &
               "			   SUM(pXiXfXe.ValorOficial) as TotalOficial," & vbCrLf &
               "			   SUM(pXiXfXe.ValorMoeda) As TotalMoeda" & vbCrLf &
               "		  FROM VW_PedidosXItensXFixacoes pXiXf" & vbCrLf &
               "		 INNER JOIN VW_PedidosXItensXFixacoesXEncargos pXiXfXe" & vbCrLf &
               "			ON pXiXfXe.Empresa_Id    = pXiXf.Empresa_Id" & vbCrLf &
               "		   AND pXiXfXe.EndEmpresa_Id = pXiXf.EndEmpresa_Id" & vbCrLf &
               "		   AND pXiXfXe.Pedido_Id     = pXiXf.Pedido_Id" & vbCrLf &
               "		   AND pXiXfXe.Produto_Id    = pXiXf.Produto_Id" & vbCrLf &
               "		   AND pXiXfXe.Fixacao_Id    = pXiXf.Fixacao_Id" & vbCrLf &
               "		 INNER JOIN #Posicao" & vbCrLf &
               "		    ON #Posicao.Empresa_Id    = pXiXf.Empresa_Id" & vbCrLf &
               "		   AND #Posicao.EndEmpresa_Id = pXiXf.EndEmpresa_Id" & vbCrLf &
               "		   AND #Posicao.Pedido_Id     = pXiXf.Pedido_Id" & vbCrLf &
               "		 WHERE pXiXf.Empresa_Id    ='" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
               "		   AND pXiXf.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
               "		   AND pXiXfXe.Encargo_Id  ='LIQUIDO'" & vbCrLf &
               "		 GROUP BY pXiXf.Empresa_Id, pXiXf.EndEmpresa_Id, pXiXf.Pedido_Id" & vbCrLf &
               "		) AS consulta" & vbCrLf &
               "  INNER JOIN #Posicao" & vbCrLf &
               "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
               "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	AND	consulta.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf


        'Atualiza Quantidade Entregue Notas Fiscais X Itens ---
        'PARTE 5 - LINHA 901
        Sql &= "Update #Posicao SET" & vbCrLf &
               "	#Posicao.Entregue           = consulta.Entregue," & vbCrLf &
               "	#Posicao.ValorLiquidoDeNota = consulta.Valor" & vbCrLf &
               "  FROM (SELECT n.Empresa_Id," & vbCrLf &
               "               n.EndEmpresa_Id," & vbCrLf &
               "               n.Pedido," & vbCrLf &
               "			   SUM(CASE" & vbCrLf &
               "			         WHEN sOP.Devolucao = 'S'" & vbCrLf &
               "			           THEN nXi.QuantidadeFiscal * - 1" & vbCrLf &
               "			           ELSE nXi.QuantidadeFiscal" & vbCrLf &
               "			       END) AS Entregue," & vbCrLf &
               "			   SUM(CASE" & vbCrLf &
               "			         WHEN sOP.Devolucao = 'S'" & vbCrLf &
               "			           THEN nxE.Valor * - 1" & vbCrLf &
               "			           ELSE nXe.Valor" & vbCrLf &
               "			       END) AS Valor" & vbCrLf &
               "		  FROM NotasFiscaisXItens nXi" & vbCrLf &
               "		 INNER JOIN NotasFiscais n" & vbCrLf &
               "			ON n.Empresa_Id      = nXi.Empresa_Id" & vbCrLf &
               "		   AND n.EndEmpresa_Id   = nXi.EndEmpresa_Id" & vbCrLf &
               "		   AND n.Cliente_Id      = nXi.Cliente_Id" & vbCrLf &
               "		   AND n.EndCliente_Id   = nXi.EndCliente_Id" & vbCrLf &
               "		   AND n.EntradaSaida_Id = nXi.EntradaSaida_Id" & vbCrLf &
               "		   AND n.Serie_Id        = nXi.Serie_Id" & vbCrLf &
               "		   AND n.Nota_Id         = nXi.Nota_Id" & vbCrLf &
               "		 INNER JOIN NotasFiscaisXEncargos nXe" & vbCrLf &
               "		 	ON nXe.Empresa_Id      = nXi.Empresa_Id" & vbCrLf &
               "		   AND nXe.EndEmpresa_Id   = nXi.EndEmpresa_Id" & vbCrLf &
               "		   AND nXe.Cliente_Id      = nXi.Cliente_Id" & vbCrLf &
               "		   AND nXe.EndCliente_Id   = nXi.EndCliente_Id" & vbCrLf &
               "		   AND nXe.EntradaSaida_Id = nXi.EntradaSaida_Id" & vbCrLf &
               "		   AND nXe.Serie_Id        = nXi.Serie_Id" & vbCrLf &
               "		   AND nXe.Nota_Id         = nXi.Nota_Id" & vbCrLf &
               "		   AND nXe.Produto_Id      = nXi.Produto_Id" & vbCrLf &
               "		   AND nXe.CFOP_Id         = nXi.CFOP_Id" & vbCrLf &
               "		   And nXe.sequencia_ID    = nXi.sequencia_ID" & vbCrLf &
               "         INNER JOIN Operacoes OP" & vbCrLf &
               "            on OP.Operacao_Id = n.Operacao" & vbCrLf &
               "		 INNER JOIN SubOperacoes sOp" & vbCrLf &
               "			ON sOp.Operacao_Id     = n.Operacao" & vbCrLf &
               "		   AND sOp.SubOperacoes_Id = n.SubOperacao" & vbCrLf &
               "		 INNER JOIN #Posicao" & vbCrLf &
               "		    ON #Posicao.Empresa_Id    = n.Empresa_Id" & vbCrLf &
               "		   AND #Posicao.EndEmpresa_Id = n.endEmpresa_Id" & vbCrLf &
               "		   AND #Posicao.Pedido_Id     = n.Pedido" & vbCrLf &
               "		 WHERE n.Situacao = 1" & vbCrLf &
               "		   AND nXe.Encargo_Id = 'LIQUIDO'" & vbCrLf &
               "		   And sOp.Classe NOT IN ('" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.FRETES.ToString & "')" & vbCrLf &
               "           and op.classe <> '" & eClassesOperacoes.FRETES.ToString & "'" & vbCrLf &
               "           And n.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf &
               "		 GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Pedido" & vbCrLf &
               "		)as consulta" & vbCrLf &
               "  INNER JOIN #Posicao" & vbCrLf &
               "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
               "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	AND	consulta.Pedido        = #Posicao.Pedido_Id" & vbCrLf


        'Atualiza Quantidade Entregue Romaneios ----
        'PARTE 6 - LINHA 935
        Sql &= "Update #Posicao SET" & vbCrLf &
               "	#Posicao.Laudo   = consulta.Laudo" & vbCrLf &
               "  from (SELECT n.Empresa_Id," & vbCrLf &
               "               n.EndEmpresa_Id," & vbCrLf &
               "               n.Pedido," & vbCrLf &
               "			   SUM(CASE" & vbCrLf &
               "			         WHEN sOp.Devolucao = 'S'" & vbCrLf &
               "			           THEN r.PesoLiquido * - 1" & vbCrLf &
               "			           ELSE r.PesoLiquido" & vbCrLf &
               "			       END) AS Laudo" & vbCrLf &
               "		  FROM NotasFiscais n" & vbCrLf &
               "		 INNER JOIN SubOperacoes sOp" & vbCrLf &
               "			ON sOp.Operacao_Id     = n.Operacao" & vbCrLf &
               "		   AND sOp.SubOperacoes_Id = n.SubOperacao" & vbCrLf &
               "		 INNER JOIN NotasFiscaisXRomaneios nXr" & vbCrLf &
               "			ON nXr.Empresa_Id      = n.Empresa_Id" & vbCrLf &
               "		   AND nXr.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
               "		   AND nXr.Cliente_Id      = n.Cliente_Id" & vbCrLf &
               "		   AND nXr.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
               "		   AND nXr.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
               "		   AND nXr.Serie_Id        = n.Serie_Id" & vbCrLf &
               "		   AND nXr.Nota_Id         = n.Nota_Id" & vbCrLf &
               "		 INNER JOIN Romaneios r" & vbCrLf &
               "			ON r.Empresa_Id    = nXr.Empresa_Id" & vbCrLf &
               "		   AND r.EndEmpresa_Id = nXr.EndEmpresa_Id" & vbCrLf &
               "		   AND r.Romaneio_Id   = nXr.Romaneio_Id" & vbCrLf &
               "		 INNER JOIN #Posicao" & vbCrLf &
               "		    ON #Posicao.Empresa_Id    = n.Empresa_Id" & vbCrLf &
               "		   AND #Posicao.EndEmpresa_Id = n.endEmpresa_Id" & vbCrLf &
               "		   AND #Posicao.Pedido_Id     = n.Pedido" & vbCrLf &
               "		 Where (n.Situacao = 1)" & vbCrLf &
               "		 GROUP BY n.Empresa_Id, n.EndEmpresa_Id, n.Pedido" & vbCrLf &
               "		) consulta" & vbCrLf &
               "  INNER JOIN #Posicao" & vbCrLf &
               "	 ON consulta.Empresa_Id    = #Posicao.Empresa_Id      " & vbCrLf &
               "	AND consulta.EndEmpresa_Id = #Posicao.EndEmpresa_Id   " & vbCrLf &
               "	AND	consulta.Pedido        = #Posicao.Pedido_Id       " & vbCrLf


        'Ajustando Saldos ----
        'PARTE 7 - LINHA 960
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "   #Posicao.Contratado = #Posicao.Entregue" & vbCrLf &
               " where #Posicao.Contratado < #Posicao.Entregue" & vbCrLf


        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "	#Posicao.AEntregar = #Posicao.Contratado" & vbCrLf &
               "						 - Case" & vbCrLf &
               "							  when #Posicao.EntradaSaida = 'E'" & vbCrLf &
               "								then Case" & vbCrLf &
               "									   when p.FreteCifFob = 'FOB'" & vbCrLf &
               "			  							 then ISNULL(#Posicao.Entregue, 0)" & vbCrLf &
               "										  else ISNULL(#Posicao.Laudo, 0)" & vbCrLf &
               "									 End" & vbCrLf &
               "								else Case" & vbCrLf &
               "									   when p.FreteCifFob = 'FOB'" & vbCrLf &
               "										 then ISNULL(#Posicao.Entregue, 0)" & vbCrLf &
               "										 else ISNULL(#Posicao.Entregue, 0)" & vbCrLf &
               "									 End" & vbCrLf &
               "						  End" & vbCrLf &
               "	FROM #Posicao" & vbCrLf &
               "   INNER JOIN Pedidos p" & vbCrLf &
               " 	  ON p.Empresa_Id    = #Posicao.Empresa_Id" & vbCrLf &
               "	 AND p.EndEmpresa_Id = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	 AND p.Pedido_Id     = #Posicao.Pedido_Id" & vbCrLf


        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "   #Posicao.AFixar = #Posicao.Contratado - isnull(#Posicao.QuantidadeFixado, 0)" & vbCrLf &
               "Delete #Posicao" & vbCrLf &
               " where #Posicao.Contratado = 0" & vbCrLf &
               "   and isnull(#Posicao.Entregue, 0) = 0" & vbCrLf


        'Atualiza Contas a Receber ----
        'PARTE 8 - LINHA 993
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "	#Posicao.Programado_ReceberOficial  = consulta.ProgramadoOficial," & vbCrLf &
               "	#Posicao.Baixado_ReceberOficial     = consulta.BaixadoOficial," & vbCrLf &
               "	#Posicao.Programado_ReceberMoeda    = consulta.ProgramadoMoeda," & vbCrLf &
               "	#Posicao.Baixado_ReceberMoeda       = consulta.BaixadoMoeda" & vbCrLf &
               " FROM (SELECT cR.EmpresaPedido," & vbCrLf &
               "              cR.EndEmpresaPedido," & vbCrLf &
               "              cR.Pedido," & vbCrLf &
               "			  ISNULL(SUM(CASE" & vbCrLf &
               "			               WHEN cR.Provisao <> 1" & vbCrLf &
               "			                 THEN cR.ValorDoDocumento" & vbCrLf &
               "			                 ELSE 0" & vbCrLf &
               "			             END), 0) AS ProgramadoOficial," & vbCrLf &
               "			  ISNULL(SUM(CASE" & vbCrLf &
               "			               WHEN cR.Provisao = 1" & vbCrLf &
               "			                 THEN cR.ValorDoDocumento" & vbCrLf &
               "			                 ELSE 0" & vbCrLf &
               "			             END), 0) AS BaixadoOficial," & vbCrLf &
               "			  ISNULL(SUM(CASE" & vbCrLf &
               "			               WHEN cR.Provisao <> 1" & vbCrLf &
               "			                 THEN cR.MoedaValorDoDocumento" & vbCrLf &
               "			                 ELSE 0" & vbCrLf &
               "			             END), 0) AS ProgramadoMoeda," & vbCrLf &
               "			  ISNULL(SUM(CASE" & vbCrLf &
               "			               WHEN cR.Provisao = 1" & vbCrLf &
               "			                 THEN cR.MoedaValorDoDocumento" & vbCrLf &
               "			                 ELSE 0" & vbCrLf &
               "			              END), 0) AS BaixadoMoeda" & vbCrLf &
               "		 FROM ContasAReceber cR" & vbCrLf &
               "		INNER JOIN #Posicao" & vbCrLf &
               "		   ON cR.EmpresaPedido    = #Posicao.Empresa_Id" & vbCrLf &
               "		  AND cR.EndEmpresaPedido = #Posicao.EndEmpresa_Id" & vbCrLf &
               "		  AND cR.Pedido           = #Posicao.Pedido_Id" & vbCrLf &
               "        INNER JOIN ComprasXProdutos CP" & vbCrLf &
               "           ON Cr.Carteira = CP.Produto_Id" & vbCrLf &
               "		WHERE cR.Situacao = 1  And CP.Adiantamento <> 'S'" & vbCrLf &
               "		and isnull(cR.Grupado,'N') <>'M'" & vbCrLf &
               "		GROUP BY cR.EmpresaPedido, cR.EndEmpresaPedido, cR.Pedido" & vbCrLf &
               "	   ) AS consulta" & vbCrLf &
               " INNER JOIN #Posicao" & vbCrLf &
               "	ON consulta.EmpresaPedido    = #Posicao.Empresa_Id" & vbCrLf &
               "   AND consulta.EndEmpresaPedido = #Posicao.EndEmpresa_Id" & vbCrLf &
               "   AND consulta.Pedido           = #Posicao.Pedido_Id" & vbCrLf


        'Atualiza Contas a Pagar ----
        'PARTE 9 - LINHA 1021
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "	#Posicao.Programado_PagarOficial = consulta.ProgramadoOficial," & vbCrLf &
               "	#Posicao.Baixado_PagarOficial    = consulta.BaixadoOficial," & vbCrLf &
               "	#Posicao.Programado_PagarMoeda   = consulta.ProgramadoMoeda," & vbCrLf &
               "	#Posicao.Baixado_PagarMoeda      = consulta.BaixadoMoeda" & vbCrLf &
               "  FROM (SELECT cP.EmpresaPedido," & vbCrLf &
               "               cP.EndEmpresaPedido," & vbCrLf &
               "               cP.Pedido," & vbCrLf &
               "			   ISNULL(SUM(CASE" & vbCrLf &
               "			                WHEN cP.Provisao <> 1" & vbCrLf &
               "			                  THEN cP.ValorDoDocumento" & vbCrLf &
               "			                  ELSE 0" & vbCrLf &
               "			              END), 0) AS ProgramadoOficial," & vbCrLf &
               "			   ISNULL(SUM(CASE" & vbCrLf &
               "			                WHEN cP.Provisao = 1" & vbCrLf &
               "			                  THEN cP.ValorDoDocumento" & vbCrLf &
               "			                  ELSE 0" & vbCrLf &
               "			              END), 0) AS BaixadoOficial," & vbCrLf &
               "			   ISNULL(SUM(CASE" & vbCrLf &
               "			                WHEN cP.Provisao <> 1" & vbCrLf &
               "			                  THEN cP.MoedaValorDoDocumento" & vbCrLf &
               "			                  ELSE 0" & vbCrLf &
               "			              END), 0) AS ProgramadoMoeda," & vbCrLf &
               "			   ISNULL(SUM(CASE" & vbCrLf &
               "			                WHEN cP.Provisao = 1" & vbCrLf &
               "			                  THEN cP.MoedaValorDoDocumento" & vbCrLf &
               "			                  ELSE 0" & vbCrLf &
               "			              END), 0) AS BaixadoMoeda" & vbCrLf &
               "		  FROM ContasAPagar cP" & vbCrLf &
               "	     INNER JOIN #Posicao" & vbCrLf &
               "		    ON cP.EmpresaPedido    = #Posicao.Empresa_Id" & vbCrLf &
               "		   AND cP.EndEmpresaPedido = #Posicao.EndEmpresa_Id" & vbCrLf &
               "		   AND cP.Pedido           = #Posicao.Pedido_Id" & vbCrLf &
               "         INNER JOIN ComprasXProdutos Cx " & vbCrLf &
               "            ON cP.Carteira = Cx.Produto_Id" & vbCrLf &
               "         WHERE cP.Situacao      = 1" & vbCrLf &
               "           And Cx.Adiantamento <> 'S'" & vbCrLf &
               "           and isnull(cp.Grupado,'N') <>'M'" & vbCrLf &
               "		 GROUP BY cP.EmpresaPedido, cP.EndEmpresaPedido, cP.Pedido" & vbCrLf &
               "		) AS consulta" & vbCrLf &
               "  INNER JOIN #Posicao" & vbCrLf &
               "	 ON consulta.EmpresaPedido    = #Posicao.Empresa_Id" & vbCrLf &
               "	AND consulta.EndEmpresaPedido = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	AND	consulta.Pedido           = #Posicao.Pedido_Id" & vbCrLf


        'Atualiza Adiantamentos de Materia Prima ----
        'PARTE 10 - LINHA 1050
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "	#Posicao.Adiantamento = consulta.Valor" & vbCrLf &
               "  FROM (SELECT cP.EmpresaPedido," & vbCrLf &
               "               cP.EndEmpresaPedido," & vbCrLf &
               "               cP.Pedido as Pedido," & vbCrLf &
               "               Sum(DebitoOficial - CreditoOficial) as Valor" & vbCrLf &
               "		  FROM ContasAPagar cP" & vbCrLf &
               "		 INNER JOIN Razao r" & vbCrLf &
               "			ON r.Titulo = cP.Registro_Id" & vbCrLf &
               "		 WHERE Conta_ID  = '1010303'" & vbCrLf &
               "		 and isnull(cP.Grupado,'N') <>'M'" & vbCrLf &
               "		 Group by cP.EmpresaPedido, cP.EndEmpresaPedido, cp.Pedido" & vbCrLf &
               "		) AS consulta" & vbCrLf &
               "  INNER JOIN #Posicao" & vbCrLf &
               "	 ON consulta.EmpresaPedido    = #Posicao.Empresa_Id" & vbCrLf &
               "	AND consulta.EndEmpresaPedido = #Posicao.EndEmpresa_Id" & vbCrLf &
               "	AND	consulta.Pedido           = #Posicao.Pedido_Id" & vbCrLf


        'Ajusta Saldo Financeiro ----------------
        'PARTE 11 - LINHA 1072
        Sql &= "UPDATE #Posicao" & vbCrLf &
               "	SET #Posicao.Programado_SaldoOficial = isnull(#Posicao.Programado_PagarOficial, 0) - isnull(#Posicao.Programado_ReceberOficial, 0)," & vbCrLf &
               "		#Posicao.Programado_SaldoMoeda   = isnull(#Posicao.Programado_PagarMoeda, 0)   - isnull(#Posicao.Programado_ReceberMoeda, 0)," & vbCrLf &
               "		#Posicao.Baixado_SaldoOficial    = isnull(#Posicao.Baixado_PagarOficial, 0)    - isnull(#Posicao.Baixado_ReceberOficial, 0)," & vbCrLf &
               "		#Posicao.Baixado_SaldoMoeda      = isnull(#Posicao.Baixado_PagarMoeda, 0)      - isnull(#Posicao.Baixado_ReceberMoeda, 0)" & vbCrLf &
               "from #Posicao" & vbCrLf


        Sql &= "Update #Posicao Set" & vbCrLf &
               "  #Posicao.Programado_SaldoOficial = (#Posicao.Programado_SaldoOficial * -1)" & vbCrLf &
               "  from #Posicao" & vbCrLf &
               " where #Posicao.Programado_SaldoOficial < 0" & vbCrLf


        Sql &= "Update #Posicao Set" & vbCrLf &
               "  #Posicao.Programado_SaldoMoeda   = (#Posicao.Programado_SaldoMoeda * -1)" & vbCrLf &
               "  from #Posicao" & vbCrLf &
               " where #Posicao.Programado_SaldoMoeda < 0" & vbCrLf


        Sql &= "Update #Posicao Set" & vbCrLf &
               "  #Posicao.Baixado_SaldoOficial = (#Posicao.Baixado_SaldoOficial * -1)" & vbCrLf &
               "  from #Posicao" & vbCrLf &
               " where #Posicao.Baixado_SaldoOficial < 0" & vbCrLf


        Sql &= "Update #Posicao Set" & vbCrLf &
               "  #Posicao.Baixado_SaldoMoeda   = (#Posicao.Baixado_SaldoMoeda * -1)" & vbCrLf &
               "  from #Posicao" & vbCrLf &
               " where #Posicao.Baixado_SaldoMoeda < 0" & vbCrLf


        'Calcula Quantidade Paga ----------------
        'PARTE 12 - LINHA 1094
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "	#Posicao.Pago = CASE" & vbCrLf &
               "					  WHEN #Posicao.MoedaPedido = 1" & vbCrLf &
               "						THEN (#Posicao.QuantidadeFixado / #Posicao.ValorFixadoOficial) * #Posicao.Baixado_SaldoOficial" & vbCrLf &
               "						ELSE (#Posicao.QuantidadeFixado / #Posicao.ValorFixadoMoeda)   * #Posicao.Baixado_SaldoMoeda" & vbCrLf &
               "					END" & vbCrLf &
               "WHERE #Posicao.ValorFixadoOficial > 0" & vbCrLf


        Sql &= "UPDATE #Posicao Set" & vbCrLf &
               "  #Posicao.RecebidoNaoPago = (#Posicao.Entregue - #Posicao.Pago)" & vbCrLf &
               " WHERE isnull(#Posicao.Entregue, 0) > #Posicao.Pago" & vbCrLf

        SqlArray.Add(Sql)

        Sql &= "UPDATE #Posicao Set" & vbCrLf &
               " #Posicao.PagoNaoRecebido = (#Posicao.Pago - isnull(#Posicao.Entregue, 0))" & vbCrLf &
               " WHERE #Posicao.Pago > isnull(#Posicao.Entregue, 0)" & vbCrLf

        'PARTE 13 - LINHA 1109
        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "   #Posicao.TotalOficial = (SELECT SUM(CASE" & vbCrLf &
               "                                         WHEN pXi.TipoDeLancamento = 'E'" & vbCrLf &
               "                                           THEN (pXi.TotalOficial * - 1)" & vbCrLf &
               "                                           ELSE pXi.TotalOficial" & vbCrLf &
               "                                       END) AS Oficial" & vbCrLf &
               "							  FROM PedidoXItemXLancamento pXi" & vbCrLf &
               "							 WHERE (pXi.Empresa_Id    = #Posicao.Empresa_Id)" & vbCrLf &
               "							   AND (pXi.EndEmpresa_Id = #Posicao.EndEmpresa_Id)" & vbCrLf &
               "							   AND (pXi.Pedido_Id     = #Posicao.Pedido_Id))," & vbCrLf &
               "	#Posicao.TotalMoeda  = (SELECT SUM(CASE" & vbCrLf &
               "	                                     WHEN pXi.TipoDeLancamento = 'E'" & vbCrLf &
               "	                                       THEN (pXi.TotalMoeda * - 1)" & vbCrLf &
               "	                                       ELSE TotalMoeda" & vbCrLf &
               "	                                    END) AS Moeda" & vbCrLf &
               "							  FROM PedidoXItemXLancamento pXi" & vbCrLf &
               "							 WHERE (pXi.Empresa_Id = #Posicao.Empresa_Id)" & vbCrLf &
               "							   AND (pXi.EndEmpresa_Id = #Posicao.EndEmpresa_Id)" & vbCrLf &
               "							   AND (pXi.Pedido_Id = #Posicao.Pedido_Id))" & vbCrLf

        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "   #Posicao.UnitarioOficial = (#Posicao.TotalOficial / #Posicao.QuantidadeFixado)" & vbCrLf &
               "  FROM #Posicao" & vbCrLf &
               " INNER JOIN Produtos p" & vbCrLf &
               "    ON p.Produto_Id = #Posicao.Produto" & vbCrLf &
               " Where #Posicao.TotalOficial     > 0" & vbCrLf &
               "   and #Posicao.QuantidadeFixado > 0" & vbCrLf

        Sql &= "UPDATE #Posicao SET" & vbCrLf &
               "  #Posicao.UnitarioMoeda = (#Posicao.TotalMoeda / #Posicao.Contratado)" & vbCrLf &
               "  FROM #Posicao" & vbCrLf &
               " INNER JOIN Produtos p" & vbCrLf &
               "	ON p.Produto_Id = #Posicao.Produto" & vbCrLf &
               " Where #Posicao.TotalMoeda > 0" & vbCrLf &
               "   and #Posicao.Contratado > 0" & vbCrLf
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("PosicaoDePedido", "RELATORIO") Then
                If Validar() Then
                    Processar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PosicaoDePedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class