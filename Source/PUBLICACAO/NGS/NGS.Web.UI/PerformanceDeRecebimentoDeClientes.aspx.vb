Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Xml
Imports System.Security
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Pkcs
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.Xml.SignedXml
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports System.Drawing

Public Class PerformanceDeRecebimentoDeClientes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("Usuarios", "ACESSAR") Then
                Limpar()
                CargaUnidadeDeNegocio()
                ddlUnidadeDeNegocio.SelectedIndex = ddlUnidadeDeNegocio.Items.IndexOf(ddlUnidadeDeNegocio.Items.FindByValue(UsuarioServidor.Usuario.AcessoUnidade))
                CargaEmpresa()
                ddlEmpresa.SelectedIndex = ddlEmpresa.Items.IndexOf(ddlEmpresa.Items.FindByValue(UsuarioServidor.Usuario.AcessoEmpresa & "-" & UsuarioServidor.Usuario.AcessoEnderecoEmpresa))
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Limpar()

        'txtData.Text = String.Format(CInt(DateTime.Now.Day) - 1 & "/" & DateTime.Now.ToString("MM/yyyy"))
        txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CargaEmpresa()
        ddlEmpresa.ClearSelection()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Function validarCampos()
        If Not ddlUnidadeDeNegocio.SelectedIndex > 0 Then
            MsgBox(Me.Page, "Unidade não selecionada.")
            Return False
        End If
        If Not ddlEmpresa.SelectedIndex > 0 Then
            MsgBox(Me.Page, "Empresa não selecionada.")
            Return False
        End If
        If txtData.Text > String.Format(CInt(DateTime.Now.Day) - 1 & "/" & DateTime.Now.ToString("MM/yyyy")) Then
            MsgBox(Me.Page, "Data selecionada não é valida.")
            Return False
        End If
        Return True
    End Function

    Private Function consultaRelatorio()
        Dim Sql As String = ""
        Dim ds As DataSet = New DataSet

        Sql = "	IF OBJECT_ID('tempdb..#tempConsultaPrincipal') IS NOT NULL" & vbCrLf & _
              " DROP TABLE #tempConsultaPrincipal;" & vbCrLf & _
              "	IF OBJECT_ID('tempdb..#tempConsultaIdentificada') IS NOT NULL" & vbCrLf & _
              "	DROP TABLE #tempConsultaIdentificada;" & vbCrLf & _
              "	SELECT Cliente," & vbCrLf & _
              "	Nota_Id AS Nota," & vbCrLf & _
              "	ValorNota," & vbCrLf & _
              "	Titulo," & vbCrLf & _
              "	DATA AS DataDaVenda," & vbCrLf & _
              "	Vencimento," & vbCrLf & _
              "	Baixa AS DataPagto," & vbCrLf & _
              "	Valor," & vbCrLf & _
              "	Pago," & vbCrLf & _
              "	(APagar - Pago) AS Saldo" & vbCrLf & _
              "	into #tempConsultaPrincipal" & vbCrLf & _
              "	FROM (" & vbCrLf & _
              "	SELECT  Empresas.Reduzido +" & vbCrLf & _
              "	' - ' + Empresas.Nome +" & vbCrLf & _
              " ' - ' + Empresas.Cidade +" & vbCrLf & _
              "	' - ' + Empresas.Estado AS Empresa," & vbCrLf & _
              "	Clientes.Nome +" & vbCrLf & _
              "	' - ' + Clientes.Cidade +" & vbCrLf & _
              "	' - ' + Clientes.Estado AS Cliente," & vbCrLf & _
              "	CAST(Nota.Nota_Id as varchar) as Nota_Id," & vbCrLf & _
              "	SUM(EncargoLiquido.Valor) as ValorNota," & vbCrLf & _
              "	convert(varchar,Nota.Movimento,103) AS DATA," & vbCrLf & _
              "	CASE" & vbCrLf & _
              "	WHEN ContasAReceber.Situacao = 1 THEN ContasAReceber.ValorDoDocumento ELSE 0 END AS Valor," & vbCrLf & _
              "	CAST(ContasAReceber.Registro_Id as varchar) as Titulo," & vbCrLf & _
              "	SubOperacoes.Operacao_Id AS OP," & vbCrLf & _
              "	SubOperacoes.SubOperacoes_Id AS SO," & vbCrLf & _
              "	convert(varchar,ContasAReceber.Vencimento,103) as Vencimento," & vbCrLf & _
              "	CASE WHEN ContasAReceber.Provisao <> 1 THEN '' ELSE convert(varchar,ContasAReceber.Baixa,103) END as Baixa," & vbCrLf & _
              "	SubOperacoes.Descricao," & vbCrLf & _
              "	ContasAReceber.ValorDoDocumento AS APagar," & vbCrLf & _
              "	CASE WHEN ContasAReceber.Provisao = 1 THEN ContasAReceber.ValorLiquido ELSE 0 END AS Pago" & vbCrLf & _
              "	FROM NotasFiscais as Nota" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens as ItensNota on ItensNota.Empresa_Id = Nota.Empresa_Id" & vbCrLf & _
              "	AND ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id" & vbCrLf & _
              "	AND ItensNota.Cliente_Id = Nota.Cliente_Id" & vbCrLf & _
              "	AND ItensNota.EndCliente_Id = Nota.EndCliente_Id" & vbCrLf & _
              "	AND ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id" & vbCrLf & _
              "	AND ItensNota.Serie_Id = Nota.Serie_Id" & vbCrLf & _
              "	AND ItensNota.Nota_Id = Nota.Nota_Id" & vbCrLf & _
              "	INNER JOIN NotasFiscaisXEncargos as EncargoLiquido on EncargoLiquido.Empresa_Id = ItensNota.Empresa_Id" & vbCrLf & _
              "	AND EncargoLiquido.EndEmpresa_Id = ItensNota.EndEmpresa_Id" & vbCrLf & _
              "	AND EncargoLiquido.Cliente_Id = ItensNota.Cliente_Id" & vbCrLf & _
              "	AND EncargoLiquido.EndCliente_Id = ItensNota.EndCliente_Id" & vbCrLf & _
              "	AND EncargoLiquido.EntradaSaida_Id = ItensNota.EntradaSaida_Id" & vbCrLf & _
              "	AND EncargoLiquido.Serie_Id = ItensNota.Serie_Id" & vbCrLf & _
              "	AND EncargoLiquido.Nota_Id = ItensNota.Nota_Id  " & vbCrLf & _
              "	AND EncargoLiquido.Produto_Id = ItensNota.Produto_Id" & vbCrLf & _
              "	AND EncargoLiquido.Sequencia_id = ItensNota.Sequencia_Id" & vbCrLf & _
              "	AND EncargoLiquido.Encargo_Id = 'LIQUIDO'" & vbCrLf & _
              "	INNER JOIN Pedidos AS p ON p.Empresa_Id = Nota.Empresa_Id" & vbCrLf & _
              "	AND p.EndEmpresa_Id = Nota.EndEmpresa_Id" & vbCrLf & _
              "	AND p.Pedido_id = Nota.Pedido" & vbCrLf & _
              "	INNER JOIN Clientes AS Empresas ON Nota.Empresa_Id = Empresas.Cliente_Id" & vbCrLf & _
              "	AND Nota.EndEmpresa_Id = Empresas.Endereco_Id" & vbCrLf & _
              "	INNER JOIN Clientes ON Nota.Cliente_Id = Clientes.Cliente_Id" & vbCrLf & _
              "	AND Nota.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " INNER JOIN SubOperacoes ON Nota.Operacao = SubOperacoes.Operacao_Id" & vbCrLf & _
              "	AND Nota.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
              "	AND SubOperacoes.Financeiro = 'S'" & vbCrLf & _
              " INNER JOIN SubOperacoes AS SubPedido ON SubPedido.Operacao_Id = p.Operacao" & vbCrLf & _
              "	AND SubPedido.SubOperacoes_Id = p.SubOperacao" & vbCrLf & _
              "	INNER JOIN NotaFiscalXTitulo ON Nota.Empresa_Id = NotaFiscalXTitulo.Empresa_Id" & vbCrLf & _
              "	AND Nota.EndEmpresa_Id = NotaFiscalXTitulo.EndEmpresa_Id" & vbCrLf & _
              "	AND Nota.Cliente_Id = NotaFiscalXTitulo.Cliente_Id" & vbCrLf & _
              "	AND Nota.EndCliente_Id = NotaFiscalXTitulo.EndCliente_Id" & vbCrLf & _
              "	AND Nota.EntradaSaida_Id = NotaFiscalXTitulo.EntradaSaida_Id" & vbCrLf & _
              "	AND Nota.Serie_Id = NotaFiscalXTitulo.Serie_Id" & vbCrLf & _
              "	AND Nota.Nota_Id = NotaFiscalXTitulo.Nota_Id" & vbCrLf & _
              "	INNER JOIN ContasAReceber ON NotaFiscalXTitulo.Titulo_Id = ContasAReceber.Registro_Id" & vbCrLf & _
              "	WHERE (SubOperacoes.EntradaSaida = 'S')" & vbCrLf & _
              " AND (SubPedido.Classe in('VENDAS','VENDASAORDEM','EXPORTACOES'))" & vbCrLf & _
              " AND (ContasAReceber.Situacao = 1)" & vbCrLf & _
              "	AND (Nota.Situacao = 1)" & vbCrLf & _
              "	AND (Nota.DataDaNota >= '" & Format(CDate(txtData.Text).AddDays(-60), "yyyy/MM/dd") & "')" & vbCrLf & _
              " AND Nota.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
              "	GROUP BY Empresas.Reduzido," & vbCrLf & _
              " Empresas.Nome, Empresas.Cidade, Empresas.Estado," & vbCrLf & _
              " Clientes.Nome, Clientes.Cidade, Clientes.Estado," & vbCrLf & _
              " Nota.Nota_Id, Nota.Movimento," & vbCrLf & _
              " ContasAReceber.Situacao, ContasAReceber.Registro_Id, ContasAReceber.Vencimento, ContasAReceber.Baixa," & vbCrLf & _
              " ContasAReceber.Provisao, ContasAReceber.ValorDoDocumento, ContasAReceber.ValorLiquido," & vbCrLf & _
              " SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao) AS Consulta" & vbCrLf & _
              " ORDER BY Cliente" & vbCrLf & _
              "" & vbCrLf

        Sql &= " SELECT ROW_NUMBER() OVER (PARTITION BY Cliente, Nota ORDER BY Cliente) As Identificador," & vbCrLf & _
               " Cliente as ClienteAgrupamento, Nota as NotaAgrupamento, * INTO #TempConsultaIdentificada FROM #tempConsultaPrincipal" & vbCrLf

        Sql &= " SELECT Cliente,Nota,ValorNota,Titulo,DataDaVenda,Vencimento,DataPagto,Valor,Pago,Saldo FROM #TempConsultaIdentificada" & vbCrLf & _
               " ORDER BY ClienteAgrupamento, NotaAgrupamento, Identificador" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Consulta")

        'CDate(txtData.Text)
        '" AND Nota.DataDaNota = '" & String.Format(CInt(DateTime.Now.Month) & "/" & DateTime.Now.ToString("dd/yyyy")) & "'" & vbCrLf & _
        Return ds
    End Function

    Private Sub gerarRelatorio()
        Dim objEmpresa As New [Lib].Negocio.Cliente(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1))

        Try
            Dim ds As DataSet = consultaRelatorio()

            If ds.Tables(0).Rows.Count < 1 Or ds Is Nothing Then
                MsgBox(Me.Page, "Sem movimento", eTitulo.Info)
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Cliente")

                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERFORMACE DE RECEBIMENTO DE CLIENTES")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'período selecionado
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DATA : " & String.Format("{0:dd/MM/yyyy}", CDate(txtData.Text)))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        If col.ColumnName <> "Row" Then
                            If col.ColumnName = "Identificador" Then
                                Continue For
                            End If
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        End If
                    Next

                    'auto filtro planilha
                    worksheet.Cells("A5:G" & rowIndex).AutoFilter = True

                    'formatação células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    'conteúdo da planilha com os dados do dataset
                    Dim Group As Integer = 1
                    Dim Nota As String = String.Empty
                    Dim Cliente As String = String.Empty

                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1

                        For Each col As DataColumn In ds.Tables(0).Columns
                            Dim Value As String = row(col.ColumnName)

                            If (row("Nota") = Nota And row("Cliente") = Cliente) And (col.ColumnName = "Nota" Or col.ColumnName = "ValorNota") Then
                                Value = String.Empty
                            End If
                            If row("Cliente") = Cliente And col.ColumnName = "Cliente" Then
                                Value = String.Empty
                            End If
                            If col.ColumnName = "Valor" Or col.ColumnName = "Saldo" Or col.ColumnName = "Pago" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToDecimal(row(col.ColumnName))
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = Value
                            End If

                            columnIndex += 1
                        Next

                        If row("Nota") <> Nota Then
                            Nota = row("Nota")
                        End If
                        If row("Cliente") <> Cliente Then
                            Cliente = row("Cliente")
                        End If

                        rowIndex += 1

                        Dim i As Integer = ds.Tables(0).Rows.IndexOf(row) + 1

                        If i = ds.Tables(0).Rows.Count Then
                        ElseIf ds.Tables(0).Rows(i).Item(0) <> row("Cliente") Then
                        ElseIf ds.Tables(0).Rows(i).Item(0) = row("Cliente") Then
                            Group += 1
                            Continue For
                        End If

                        'criando colunas de totalizadores
                        worksheet.Cells(rowIndex, 7).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, 7).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(rowIndex, 7).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(rowIndex, 7).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(rowIndex, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, 7).Value = String.Format("TOTAL")

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Formula = String.Format("=SUM(H{0}:H{1})", rowIndex - Group, rowIndex - 1)
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("I{0}", rowIndex)).Formula = String.Format("=SUM(I{0}:I{1})", rowIndex - Group, rowIndex - 1)
                        worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(J{0}:J{1})", rowIndex - Group, rowIndex - 1)
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "0.00_ ;[Red]-0.00"

                        Group = 1

                        rowIndex += 2
                    Next

                    rowIndex += 1

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        CargaUnidadeDeNegocio()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Usuarios", "RELATORIO") Then
                If validarCampos() Then
                    gerarRelatorio()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Usuarios")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class