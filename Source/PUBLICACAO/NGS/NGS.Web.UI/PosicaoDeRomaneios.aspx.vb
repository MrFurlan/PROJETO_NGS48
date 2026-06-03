Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PosicaoDeRomaneios
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.ProducaoEstoque)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PosicaoDeRomaneios", "ACESSAR") Then
                    txtDataInicial.Text = String.Format("{0:dd/MM/yyyy}", DateTime.Now)
                    txtDataFinal.Text = String.Format("{0:dd/MM/yyyy}", DateTime.Now)
                    CarregarUnidade()
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa, True)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataSet(ByVal objEmpresa As [Lib].Negocio.Cliente) As DataSet
        Sql = "SELECT RO.Empresa_Id, RO.EndEmpresa_Id, RO.Pedido, Ped.Cliente, Ped.EndCliente, CliXP.Nome, " & vbCrLf &
              " RO.Produto, RO.Operacao, RO.SubOperacao, RO.Romaneio_Id, CONVERT(VARCHAR, RO.Movimento, 103) AS Movimento, RO.PesoLiquido, " & vbCrLf &
              " isnull(RO.Processo,'') AS Processo, isnull(RxP.Pesagem_Id,0) AS Pesagem, " & vbCrLf &
              " CASE " & vbCrLf &
              " WHEN ISNULL(NxI.Nota_Id,0) > 0" & vbCrLf &
              " THEN NxI.EntradaSaida_Id + '-' + CAST(NxI.Nota_Id AS VARCHAR) + '-' + NxI.Serie_Id " & vbCrLf &
              " ELSE '0'" & vbCrLf &
              " END AS Nota_Id," & vbCrLf &
              " CASE" & vbCrLf &
              " WHEN ISNULL(NotasFiscais.Movimento,'1900') = '1900'" & vbCrLf &
              " THEN ''" & vbCrLf &
              " ELSE CONVERT(VARCHAR, NotasFiscais.Movimento, 103)" & vbCrLf &
              " END AS DataNota, " & vbCrLf &
              " NxI.QuantidadeFisica, isnull(NxI.Deposito,'') AS DepositoNota, " & vbCrLf &
              " isnull(NxI.EndDeposito,0) AS EndDepositoNota, isnull(DNF.Nome,'') AS NomeDeposito" & vbCrLf &
              " FROM SubOperacoes" & vbCrLf &
              " INNER JOIN Romaneios RO" & vbCrLf &
              " ON SubOperacoes.Operacao_Id     = RO.Operacao" & vbCrLf &
              " AND SubOperacoes.SubOperacoes_Id = RO.SubOperacao" & vbCrLf &
              " INNER JOIN Produtos Prod                           " & vbCrLf &
              " ON RO.Produto  = Prod.Produto_Id             " & vbCrLf &
              " LEFT JOIN NotasFiscaisXRomaneios NxR" & vbCrLf &
              " ON RO.Empresa_Id    = NxR.Empresa_Id " & vbCrLf &
              " AND RO.EndEmpresa_Id = NxR.EndEmpresa_Id " & vbCrLf &
              " AND RO.Romaneio_Id   = NxR.Romaneio_Id" & vbCrLf &
              " LEFT JOIN NotasFiscaisxItens NxI" & vbCrLf &
              " ON NxR.Empresa_Id      = NxI.Empresa_Id  " & vbCrLf &
              " AND NxR.EndEmpresa_Id   = NxI.EndEmpresa_Id  " & vbCrLf &
              " AND NxR.Cliente_Id      = NxI.Cliente_Id  " & vbCrLf &
              " AND NxR.EndCliente_Id   = NxI.EndCliente_Id  " & vbCrLf &
              " AND NxR.EntradaSaida_Id = NxI.EntradaSaida_Id  " & vbCrLf &
              " AND NxR.Serie_Id        = NxI.Serie_Id  " & vbCrLf &
              " AND NxR.Nota_Id         = NxI.Nota_Id  " & vbCrLf &
              " LEFT JOIN NotasFiscais " & vbCrLf &
              " ON NxI.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
              " AND NxI.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
              " AND NxI.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
              " AND NxI.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
              " AND NxI.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
              " AND NxI.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
              " AND NxI.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
              " LEFT JOIN Clientes DNF                           " & vbCrLf &
              " ON NxI.Deposito    = DNF.Cliente_Id" & vbCrLf &
              " AND NxI.EndDeposito = DNF.Endereco_Id            " & vbCrLf &
              " LEFT JOIN Pedidos Ped                           " & vbCrLf &
              " ON Ped.Empresa_Id    = RO.Empresa_Id " & vbCrLf &
              " AND Ped.EndEmpresa_Id = RO.EndEmpresa_Id " & vbCrLf &
              " AND Ped.Pedido_Id     = RO.Pedido" & vbCrLf &
              " LEFT JOIN Clientes CliXP                           " & vbCrLf &
              " ON Ped.Cliente    = CliXP.Cliente_Id" & vbCrLf &
              " AND Ped.EndCliente = CliXP.Endereco_Id            " & vbCrLf &
              " LEFT JOIN RomaneiosXPesagens RxP " & vbCrLf &
              " ON RO.Empresa_Id    = RxP.Empresa_Id " & vbCrLf &
              " AND RO.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf &
              " AND RO.Romaneio_Id   = RxP.Romaneio_Id" & vbCrLf &
              " WHERE (SubOperacoes.EstoqueFisico = 'S')" & vbCrLf &
              " AND (RO.EntradaSaida = '" & IIf(RadEntradas.Checked, "E", "S") & "')" & vbCrLf
        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Emp() As String = ddlEmpresa.SelectedValue.Split("-")
            Sql &= " AND (RO.Empresa_Id = '" & Emp(0) & "')"
            Sql &= " AND (RO.EndEmpresa_Id = " & Emp(1) & ")"
        End If

        Sql &= " AND (RO.Movimento BETWEEN '" & String.Format("{0:yyyy/MM/dd}", CDate(txtDataInicial.Text)) & "' AND '" & String.Format("{0:yyyy/MM/dd}", CDate(txtDataFinal.Text)) & "')"

        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList
            RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "RO.Produto", "")
            Sql &= " AND " & RetornoProdutos(0)
        End If

        Sql &= " ORDER BY RO.Empresa_Id, RO.EndEmpresa_Id, RO.Produto, RO.Movimento, RO.Romaneio_Id"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "PosicaoDeRomaneios")

        Return ds
    End Function

    Private Function FileToByteArray(ByVal fileName As String) As Byte()
        Dim buffer As Byte() = Nothing

        Try
            Dim stream As New System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            Dim reader As New System.IO.BinaryReader(stream)
            Dim bytes As Long = New System.IO.FileInfo(fileName).Length
            buffer = reader.ReadBytes(CType(bytes, Long))
            stream.Close()
            stream.Dispose()
            reader.Close()
        Catch ex As Exception
            Debug.WriteLine("Exception caught in process: {0}", ex.ToString())
            Throw ex
        End Try

        Return buffer
    End Function

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub Limpar()
        ucSelecaoProduto.Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CarregarEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
    End Sub

    Private Sub GerarExcel()
        Try
            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue.ToString) Then
                Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                Sql = " SELECT Cliente_Id, Endereco_Id " & vbCrLf &
                      " FROM   Clientes LEFT OUTER JOIN" & vbCrLf &
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                      " WHERE  Clientes.Cliente_Id = '" & strEmpresa(0) & "'" & vbCrLf

                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
                Dim objEmpresa As New [Lib].Negocio.Cliente(ds.Tables(0).Rows(0).Item("Cliente_Id"), ds.Tables(0).Rows(0).Item("Endereco_Id"))
                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                ds = getDataSet(objEmpresa)

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando planilha títulos
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("POSIÇÃO DE ROMANEIOS")

                        'criando linha com o cabeçalho da planilha
                        Dim rowIndex As Integer = 1
                        Dim columnIndex As Integer = 1

                        'criando linha que informa o nome da empresa e o cnpj
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa a cidade e o estado da empresa
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o título do relatório
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO DE ROMANEIOS")
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)) & " - " & IIf(RadEntradas.Checked, "ENTRADAS", "SAÍDAS"))
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:T" & rowIndex).AutoFilter = True

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using
                        rowIndex += 1

                        'criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                columnIndex += 1
                            Next

                            'aplicando formatação nas células do conteúdo
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
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

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            GerarExcel()
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
            Funcoes.Ajuda(Me.Page, "PosicaoDeRomaneios")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class