Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Imports System.Drawing
Public Class ucTransferencias
    Inherits BaseUserControl

    Dim Sql As String
    Private ds As DataSet


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub InicializarUC(ByVal cEmpresa As String, ByVal cEndEmpresa As Integer)
        Session("ssEmpresa") = cEmpresa
        Session("ssEndEmpresa") = cEndEmpresa
        CargaNotasXItens(cEmpresa, cEndEmpresa)
    End Sub

    Private Sub CargaNotasXItens(ByVal cEmpresa As String, ByVal cEndEmpresa As Integer)

        Sql = " SELECT DISTINCT  " & vbCrLf &
              "        TRY_CAST(convert(varchar,GETDATE(),23) AS DATETIME) AS Movimento, NF.Movimento AS DataNota, " & vbCrLf &
              "        NF.UsuarioInclusaoData, " & vbCrLf &
              "        NF.Empresa_Id, " & vbCrLf &
              "        NF.EndEmpresa_Id, " & vbCrLf &
              "        NF.Cliente_Id AS Cliente, " & vbCrLf &
              "        NF.EndCliente_Id AS EndCliente, " & vbCrLf &
              "        C.Nome as NomeCliente, " & vbCrLf &
              "        NF.EntradaSaida_Id AS ES, " & vbCrLf &
              "        NF.Serie_Id AS Serie, " & vbCrLf &
              "        NF.Nota_Id AS Nota, " & vbCrLf &
              "        NF.Operacao, " & vbCrLf &
              "        NF.SubOperacao, " & vbCrLf &
              "        SUM(NFxI.QuantidadeFiscal) AS Quantidade, " & vbCrLf &
              "        SUM(NFxI.Unitario) AS Unitario, " & vbCrLf &
              "        SUM(NFxI.Valor) AS Valor " & vbCrLf &
              "   FROM NotasFiscais NF " & vbCrLf &
              "  INNER JOIN NotasFiscaisxItens NFxI " & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
              "  INNER JOIN NFeRealizadas NFR " & vbCrLf &
              "     ON NFR.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
              "    AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
              "    AND NFR.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
              "    AND NFR.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
              "    AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
              "    AND NFR.Serie_Id        = NF.Serie_Id " & vbCrLf &
              "    AND NFR.Nota_Id         = NF.Nota_Id " & vbCrLf &
              "  INNER JOIN Clientes C " & vbCrLf &
              "     ON C.Cliente_Id  = NF.Cliente_Id " & vbCrLf &
              "    AND C.Endereco_Id = NF.EndCliente_Id " & vbCrLf &
              "  INNER JOIN SubOperacoes SO " & vbCrLf &
              "     ON SO.Operacao_Id     = NF.Operacao " & vbCrLf &
              "    AND SO.SubOperacoes_Id = NF.SubOperacao " & vbCrLf &
              "  LEFT JOIN DocumentoXML DXml  " & vbCrLf &
              "     ON DXml.Chave_Id  = NFR.ChaveNfe " & vbCrLf &
              "    AND DXml.Situacao = 104 " & vbCrLf &
              "  LEFT JOIN NotasFiscais NFENT  " & vbCrLf &
              "     ON NFENT.Empresa_Id      = NF.Cliente_Id  " & vbCrLf &
              "    AND NFENT.EndEmpresa_Id   = NF.EndCliente_Id " & vbCrLf &
              "    AND NFENT.Cliente_Id      = NF.Empresa_Id " & vbCrLf &
              "    AND NFENT.EndCliente_Id   = NF.EndEmpresa_Id " & vbCrLf &
              "    AND NFENT.EntradaSaida_Id = 'E' " & vbCrLf &
              "    AND NFENT.Serie_Id        = NF.Serie_Id " & vbCrLf &
              "    AND NFENT.Nota_Id         = NF.Nota_Id " & vbCrLf &
              "  WHERE NF.Cliente_Id    = '" & cEmpresa & "'" & vbCrLf &
              "    AND NF.EndCliente_Id = " & cEndEmpresa & vbCrLf &
              "    AND NF.Eletronica    = 'S'" & vbCrLf &
              "    AND (NF.NossaEmissao  = 'S' OR isnull(DXml.Situacao,0) = 104)" & vbCrLf &
              "    AND NF.Situacao in(1) " & vbCrLf &
              "    AND NF.TipoDeDocumento in(1) " & vbCrLf &
              "    AND SO.CLASSE = 'TRANSFERENCIAS'" & vbCrLf &
              "    AND ISNULL(NFENT.Nota_Id, 0) = 0" & vbCrLf


        Sql &= "  GROUP BY NF.Movimento, " & vbCrLf &
               "        NF.UsuarioInclusaoData, " & vbCrLf &
               "        NF.Empresa_Id, " & vbCrLf &
               "        NF.EndEmpresa_Id, " & vbCrLf &
               "        NF.Cliente_Id, " & vbCrLf &
               "        NF.EndCliente_Id, " & vbCrLf &
               "        C.Nome, " & vbCrLf &
               "        NF.EntradaSaida_Id, " & vbCrLf &
               "        NF.Serie_Id, " & vbCrLf &
               "        NF.Nota_Id, " & vbCrLf &
               "        NF.Operacao, " & vbCrLf &
               "        NF.SubOperacao " & vbCrLf &
               "  ORDER BY NF.UsuarioInclusaoData DESC, ES, Serie, Nota" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Notas")
        GridNotas.DataSource = ds
        GridNotas.DataBind()
    End Sub

    Protected Sub GridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridNotas.SelectedIndexChanged
        Try
            Dim Sqls As New ArrayList
            Dim nfOrigem As New NotaFiscal

            If CDate(CType(GridNotas.Rows(GridNotas.SelectedIndex).FindControl("txtMovimento"), TextBox).Text) > Today() Then
                MsgBox(Me.Page, "Data do Movimento não pode ser maior que a data atual!")
                Exit Sub
            End If

            If CDate(CType(GridNotas.Rows(GridNotas.SelectedIndex).FindControl("txtMovimento"), TextBox).Text) < CDate(GridNotas.SelectedRow.Cells(2).Text()) Then
                MsgBox(Me.Page, "Data do Movimento não pode ser menor que a data da Nota!")
                Exit Sub
            End If

            nfOrigem.CodigoEmpresa = GridNotas.SelectedRow.Cells(3).Text()
            nfOrigem.EnderecoEmpresa = GridNotas.SelectedRow.Cells(4).Text()
            nfOrigem.CodigoCliente = GridNotas.SelectedRow.Cells(5).Text()
            nfOrigem.EnderecoCliente = GridNotas.SelectedRow.Cells(6).Text()
            nfOrigem.EntradaSaida = If(GridNotas.SelectedRow.Cells(8).Text() = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
            nfOrigem.Serie = GridNotas.SelectedRow.Cells(9).Text()
            nfOrigem.Codigo = GridNotas.SelectedRow.Cells(10).Text()
            nfOrigem = New NotaFiscal(nfOrigem)

            Sql = "SELECT Operacao_Id, SubOperacao_Id, OperacaoDestino_Id, SubOperacaoDestino_Id" & vbCrLf &
                    "FROM DeParaOperacao" & vbCrLf &
                    "WHERE Operacao_Id    = " & nfOrigem.CodigoOperacao & vbCrLf &
                    "  AND SubOperacao_Id = " & nfOrigem.CodigoSubOperacao

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "DeParaOperacao")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count = 1 Then

                Sql = "update Pedidos " & vbCrLf &
                      "       Set FiscalAberto = 0, FinanceiroAberto = 0 " & vbCrLf &
                      "WHERE Empresa_Id = '" & nfOrigem.CodigoEmpresa & "'" & vbCrLf &
                      "  AND EndEmpresa_Id = " & nfOrigem.EnderecoEmpresa & vbCrLf &
                      "  AND Pedido_Id = " & nfOrigem.CodigoPedido

                Sqls.Add(Sql)

                Sql = String.Empty

                Dim objNota = nfOrigem.Codigo
                Dim objSerie = nfOrigem.Serie

                nfOrigem.IUD = "I"
                nfOrigem.NossaEmissao = False
                nfOrigem.NotaProdutor = objNota
                nfOrigem.SerieNotaProdutor = objSerie
                nfOrigem.Movimento = CDate(CType(GridNotas.Rows(GridNotas.SelectedIndex).FindControl("txtMovimento"), TextBox).Text)

                'Faz a troca da Operação
                nfOrigem.CodigoOperacao = ds.Tables(0).Rows(0).Item("OperacaoDestino_Id")
                nfOrigem.CodigoSubOperacao = ds.Tables(0).Rows(0).Item("SubOperacaoDestino_Id")

                'Troca a NF de Saída para Entrada
                nfOrigem.EntradaSaida = eEntradaSaida.Entrada

                'Faz a troca entre a Filial de Saída e Entrada
                nfOrigem.CodigoEmpresa = GridNotas.SelectedRow.Cells(5).Text()
                nfOrigem.EnderecoEmpresa = GridNotas.SelectedRow.Cells(6).Text()
                nfOrigem.Empresa = New Cliente(nfOrigem.CodigoEmpresa, nfOrigem.EnderecoEmpresa)

                nfOrigem.CodigoCliente = GridNotas.SelectedRow.Cells(3).Text()
                nfOrigem.EnderecoCliente = GridNotas.SelectedRow.Cells(4).Text()
                nfOrigem.Cliente = New Cliente(nfOrigem.CodigoCliente, nfOrigem.EnderecoCliente)

                nfOrigem.CodigoDeposito = GridNotas.SelectedRow.Cells(5).Text()
                nfOrigem.EnderecoDeposito = GridNotas.SelectedRow.Cells(6).Text()
                nfOrigem.Deposito = New Cliente(nfOrigem.CodigoEmpresa, nfOrigem.EnderecoEmpresa)

                nfOrigem.CodigoDestino = GridNotas.SelectedRow.Cells(3).Text()
                nfOrigem.EnderecoDestino = GridNotas.SelectedRow.Cells(4).Text()
                nfOrigem.Destino = New Cliente(nfOrigem.CodigoCliente, nfOrigem.EnderecoCliente)

                nfOrigem.CFOP = Nothing

                For Each objItemNF In nfOrigem.Itens
                    objItemNF.Encargos.Clear()

                    objItemNF.CodigoDeposito = GridNotas.SelectedRow.Cells(5).Text()
                    objItemNF.EnderecoDeposito = GridNotas.SelectedRow.Cells(6).Text()
                    objItemNF.Deposito = New Cliente(nfOrigem.CodigoEmpresa, nfOrigem.EnderecoEmpresa)

                    objItemNF.CodigoOperacao = nfOrigem.CodigoOperacao
                    objItemNF.CodigoSubOperacao = nfOrigem.CodigoSubOperacao
                    objItemNF.CFOP = 0
                    objItemNF.CodigoOperacaoEstado = 0

                    'Busca os encargos para Entrada
                    objItemNF.CarregandoEncargos = True
                    objItemNF.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                    objItemNF.CarregandoEncargos = False

                    If objItemNF.Encargos.Count = 0 Then
                        MsgBox(Me.Page, "Encargos da operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " não foram encontrados na filial " & nfOrigem.Empresa.Codigo & "-" & nfOrigem.Empresa.Cidade & "/" & nfOrigem.Empresa.CodigoEstado)
                        Exit Sub
                    End If
                Next

                'Cria o Pedido
                nfOrigem.CodigoPedido = 0
                Dim objPedido As New Pedido(nfOrigem, 1)

                'Busca o numerador do Pedido
                Dim SqlN As String = "exec sp_Numerador '" & objPedido.CodigoEmpresa & "'," & objPedido.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
                Dim dsN As New DataSet
                dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                Dim CodigoNumerador As Integer = 0
                If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                    CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                End If

                If Not CodigoNumerador > 0 Then
                    MsgBox(Me.Page, "Numerador de Pedidos não cadastrado!")
                    Exit Sub
                End If

                objPedido.Codigo = CodigoNumerador
                objPedido.FiscalAberto = False
                objPedido.FinanceiroAberto = False
                nfOrigem.CodigoPedido = CodigoNumerador
                nfOrigem.Pedido = objPedido
                nfOrigem.CIFFOB = objPedido.FreteCIFFOB
                nfOrigem.UsuarioInclusao = Session("ssNomeUsuario")
                nfOrigem.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")

                If nfOrigem.Romaneio IsNot Nothing AndAlso nfOrigem.Romaneio.Codigo > 0 Then
                    nfOrigem.CriarRomaneio = True
                End If

                If RealizarManifestoNFE(nfOrigem) Then
                    objPedido.SalvarSql(Sqls)
                    nfOrigem.SalvarSql(Sqls)

                    If Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, "Transferência incluída com sucesso!")

                        CargaNotasXItens(GridNotas.SelectedRow.Cells(5).Text(), GridNotas.SelectedRow.Cells(6).Text())

                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If

            Else
                MsgBox(Me.Page, "Operação de Transferência não foi encontrada no DEPARA. Entre em contato com o Suporte.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub


    Private Function RealizarManifestoNFE(ByRef objNotaFiscal As NotaFiscal) As Boolean

        Dim msgResultado As String = String.Empty

        Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)

        'Realiza o manifesto da NFe
        If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso ModeloNFe.Equals("55") Then '(Modelo: 55 NFe)

            ''Download do Arquivo.
            Dim bytes As Byte() = New FilesManager().getFileXml(String.Format("{0}-nfe.xml", objNotaFiscal.ChaveNFE))
            If bytes Is Nothing Then
                MsgBox(Me.Page, "XML não foi encontrado, favor inserir o manualmente.")
                Return False
            Else
                Dim caminhoArquivoFile As String = Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE))
                If Not File.Exists(caminhoArquivoFile) Then
                    System.IO.File.WriteAllBytes(caminhoArquivoFile, bytes)
                End If
            End If

            If Not DocumentoEletronico.ManifestoNFe(objNotaFiscal, eTipoManifesto.ConfirmacaoDaOperacao, msgResultado) Then
                MsgBox(Me.Page, msgResultado)
                Return False
            End If

            If bytes IsNot Nothing Then
                Try
                    If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                        Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE))
                        If File.Exists(caminhoArquivo) Then
                            File.Delete(caminhoArquivo)
                        End If
                        System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                    End If

                    'Leitura do Arquivo.
                    Dim DsXml As New DataSet
                    DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE)))

                    If Not objNotaFiscal.CodigoEmpresa = DsXml.Tables("dest").Rows(0)("CNPJ").ToString() Then
                        MsgBox(Me.Page, "Empresa do XML está diferente da informanda na Nota Fiscal, verifique. Se o erro persistir entre em contato com o suporte.", eTitulo.Erro)
                        Return False
                    End If

                    Dim temArquivo As Boolean = False

                    If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                        For Each arq In objNotaFiscal.Arquivos
                            If arq.Descricao = String.Format("{0}-nfe.xml", objNotaFiscal.ChaveNFE) Then temArquivo = True
                        Next
                    End If

                    If Not temArquivo Then
                        objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                                             .IUD = "I",
                                             .Codigo = String.Empty,
                                             .Descricao = String.Format("{0}-nfe.xml", objNotaFiscal.ChaveNFE),
                                             .Arquivo = bytes})
                    End If

                    Return True

                Catch ex As Exception
                    Throw New Exception(ex.Message)
                    Return False
                End Try
            Else
                MsgBox(Me.Page, "XML não encontrado.")
                Return False
            End If
        Else
            If Not ModeloNFe.Equals("55") Then
                MsgBox(Me.Page, "Manifesto permitido somente para Nota Fiscal")
                Return False
            End If
        End If

    End Function

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorioDados() 'Excel Dados
    End Sub

    Private Sub EmitirRelatorioDados()
        Try
            Dim t = Session("empresa")
            Dim tt = Session("endEmpresa")

            Dim objEmpresa As New [Lib].Negocio.Cliente(t, tt)
            'Dim data As String = CStr(txtDataInicial.Text & " - " & txtDataFinal.Text)

            ds = getDataset()

            Dim dt As DataTable = New DataTable()

            dt = ds.Tables(0)

            If dt.Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando planilha e título
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório De transferêcia de notas.")

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
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "TRANSFERÊNCIA DE NOTAS")
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    ' criando cabeçalho da planilha com os dados do dataset
                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'formatando células numéricas
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000" ' quantidade
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,########0.00000000_ ;[Red]-#,########0.00000000" ' unitário
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00" ' valor

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dt.Rows
                        columnIndex = 1

                        For Each col As DataColumn In dt.Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("A{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataset() As DataSet
        Sql = " SELECT DISTINCT  " & vbCrLf &
              "        TRY_CAST(convert(varchar,GETDATE(),23) AS DATETIME) AS Movimento, NF.Movimento AS DataNota, " & vbCrLf &
              "        NF.Empresa_Id,                                                                              " & vbCrLf &
              "        NF.EndEmpresa_Id,                                                                           " & vbCrLf &
              "        NF.Cliente_Id AS Cliente,                                                                   " & vbCrLf &
              "        NF.EndCliente_Id AS EndCliente,                                                             " & vbCrLf &
              "        C.Nome as NomeCliente,                                                                      " & vbCrLf &
              "        NF.EntradaSaida_Id AS ES,                                                                   " & vbCrLf &
              "        NF.Serie_Id AS Serie,                                                                       " & vbCrLf &
              "        NF.Nota_Id AS Nota,                                                                         " & vbCrLf &
              "        NF.Operacao,                                                                                " & vbCrLf &
              "        NF.SubOperacao,                                                                             " & vbCrLf &
              "		NFREa.ChaveNfe,                                                                                " & vbCrLf &
              "        SUM(NFxI.QuantidadeFiscal) AS Quantidade,                                                   " & vbCrLf &
              "        SUM(NFxI.Unitario) AS Unitario,                                                             " & vbCrLf &
              "        SUM(NFxI.Valor) AS Valor,                                                                   " & vbCrLf &
              "        NF.UsuarioInclusao,                                                                         " & vbCrLf &
              "        NF.UsuarioInclusaoData                                                                      " & vbCrLf &
              "   FROM NotasFiscais NF                                                                             " & vbCrLf &
              "  INNER JOIN NotasFiscaisxItens NFxI                                                                " & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id                                                        " & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id                                                     " & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id                                                        " & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id                                                     " & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id                                                   " & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id                                                          " & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id                                                           " & vbCrLf &
              "  LEFT JOIN NFERealizadas NFREa                                                                     " & vbCrLf &
              "     ON NFREa.Empresa_Id      = NF.Empresa_Id                                                       " & vbCrLf &
              "    AND NFREa.EndEmpresa_Id   = NF.EndEmpresa_Id                                                    " & vbCrLf &
              "    AND NFREa.Cliente_Id      = NF.Cliente_Id                                                       " & vbCrLf &
              "    AND NFREa.EndCliente_Id   = NF.EndCliente_Id                                                    " & vbCrLf &
              "    AND NFREa.EntradaSaida_Id = NF.EntradaSaida_Id                                                  " & vbCrLf &
              "    AND NFREa.Serie_Id        = NF.Serie_Id                                                         " & vbCrLf &
              "    AND NFREa.Nota_Id         = NF.Nota_Id                                                          " & vbCrLf &
              "  INNER JOIN Clientes C                                                                             " & vbCrLf &
              "     ON C.Cliente_Id  = NF.Cliente_Id                                                               " & vbCrLf &
              "    AND C.Endereco_Id = NF.EndCliente_Id                                                            " & vbCrLf &
              "  INNER JOIN SubOperacoes SO                                                                        " & vbCrLf &
              "     ON SO.Operacao_Id     = NF.Operacao                                                            " & vbCrLf &
              "    AND SO.SubOperacoes_Id = NF.SubOperacao                                                         " & vbCrLf &
              "  LEFT JOIN NotasFiscais NFENT                                                                      " & vbCrLf &
              "     ON NFENT.Empresa_Id      = NF.Cliente_Id                                                       " & vbCrLf &
              "    AND NFENT.EndEmpresa_Id   = NF.EndCliente_Id                                                    " & vbCrLf &
              "    AND NFENT.Cliente_Id      = NF.Empresa_Id                                                       " & vbCrLf &
              "    AND NFENT.EndCliente_Id   = NF.EndEmpresa_Id                                                    " & vbCrLf &
              "    AND NFENT.EntradaSaida_Id = 'E'                                                                 " & vbCrLf &
              "    AND NFENT.Serie_Id        = NF.Serie_Id                                                         " & vbCrLf &
              "    AND NFENT.Nota_Id         = NF.Nota_Id                                                          " & vbCrLf &
              "  WHERE NF.Eletronica    = 'S'                                                                      " & vbCrLf &
              "    AND NF.NossaEmissao  = 'S'                                                                      " & vbCrLf &
              "    AND NF.Situacao in(1)                                                                           " & vbCrLf &
              "    AND NF.TipoDeDocumento in(1)                                                                    " & vbCrLf &
              "    AND SO.CLASSE = 'TRANSFERENCIAS'                                                                " & vbCrLf &
              "    AND ISNULL(NFENT.Nota_Id, 0) = 0                                                                " & vbCrLf &
              "    AND NFREa.ChaveNfe NOT IN('21240724450490000609550100000000131241225873','21240724450490000609550050000000061884338176')" & vbCrLf &
              "  GROUP BY NF.Movimento,                                                                            " & vbCrLf &
              "        NF.Empresa_Id,                                                                              " & vbCrLf &
              "        NF.EndEmpresa_Id,                                                                           " & vbCrLf &
              "        NF.Cliente_Id,                                                                              " & vbCrLf &
              "        NF.EndCliente_Id,                                                                           " & vbCrLf &
              "        C.Nome,                                                                                     " & vbCrLf &
              "        NF.EntradaSaida_Id,                                                                         " & vbCrLf &
              "        NF.Serie_Id,                                                                                " & vbCrLf &
              "        NF.Nota_Id,                                                                                 " & vbCrLf &
              "        NF.Operacao,                                                                                " & vbCrLf &
              "        NF.SubOperacao,                                                                             " & vbCrLf &
              "		NFREa.ChaveNfe,                                                                                " & vbCrLf &
              "        NF.UsuarioInclusao,                                                                         " & vbCrLf &
              "        NF.UsuarioInclusaoData                                                                      " & vbCrLf &
              "  ORDER BY NF.Empresa_Id DESC, ES, Serie, Nota                                                      " & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "NotasFiscais")

        Return ds
    End Function

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divTransferencias")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class