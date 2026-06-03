Imports System.Drawing
Imports System.IO
Imports NGS.Lib.Negocio
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Public Class RelatorioMovimentoFiscalPorOperacao
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioMovimentoFiscalPorOperacao", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, [Lib].Negocio.CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
                    ddl.Carregar(ddlClasse, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1")
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
                    txtData1.Text = String.Format("01/{0}/{1}", Now.Month().ToString("00"), Now.Year())
                    txtData2.Text = Now.ToString("dd/MM/yyyy")
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página. É necessário ter as permissões ACESSAR e RELATORIO no Processo RelatorioMovimentoFiscalPorOperacao", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioMovimentoFiscalPorOperacao", "Relatorio") Then
                If validaCampos() Then
                    Dim ds As DataSet = getDataSet()
                    Dim param As New Dictionary(Of String, Object)
                    getParametrosConsulta(param)
                    Funcoes.BindReport(Me.Page, ds, "Cr_MovimentoFiscalPorOperacao", eExportType.PDF, param)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioMovimentoFiscalPorOperacao", "Relatorio") Then
                If validaCampos() Then

                    Dim ds As DataSet = getDataSet()

                    ds.Tables(0).Columns.Remove("EmpresaReduzido")
                    ds.Tables(0).Columns.Remove("EndEmpresa_Id")
                    ds.Tables(0).Columns.Remove("DescGrupo")

                    Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    If File.Exists(fileName) Then
                        File.Delete(fileName)
                    End If

                    Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))

                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)

                            'criando planilha títulos
                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("NOTAS")


                            'criando linha com o cabeçalho da planilha
                            Dim rowIndex As Integer = 1
                            Dim columnIndex As Integer = 1

                            'criando linha que informa o nome da empresa e o cnpj
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa a cidade e o estado da empresa
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa o título do relatório
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Relatório de Movimentações Fiscais por Operação.")
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa o período selecionado na página
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtData1.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtData2.Text)))
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1


                            'criando linha com o cabeçalho da planilha
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            Next

                            'criando auto filtro na planilha
                            worksheet.Cells("A5:N" & rowIndex).AutoFilter = True

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using
                            rowIndex += 1

                            For Each row As DataRow In ds.Tables(0).Rows
                                columnIndex = 1
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    columnIndex += 1
                                Next

                                worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

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

                            'congelando primeira linham
                            worksheet.View.FreezePanes(2, 1)

                            'salvando planilha do excel
                            package.Save()
                        End Using
                    End Using

                    'download do arquivo pelo browser
                    Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                Exit Sub
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
            Funcoes.Ajuda(Me.Page, "RelatorioFiscalPorOperacao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaCliente_Click(sender As Object, e As EventArgs) Handles btnBuscaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"
    Private Function validaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlClasse.SelectedValue) Then
            MsgBox(Me.Page, "Informe a Classe.")
            Return False
        ElseIf (String.IsNullOrWhiteSpace(txtData1.Text) OrElse String.IsNullOrWhiteSpace(txtData2.Text)) AndAlso String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            MsgBox(Me.Page, "Informe um período ou safra.")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        ddlClasse.SelectedValue = String.Empty
        ddlSafra.SelectedValue = String.Empty
        txtData1.Text = String.Format("01/{0}/{1}", Now.Month().ToString("00"), Now.Year())
        txtData2.Text = Now.ToString("dd/MM/yyyy")
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "   SELECT E.Reduzido AS EmpresaReduzido," & vbCrLf &
                            "          dbo.FormatarCpfCnpj(NF.Empresa_Id) as Empresa_Id," & vbCrLf &
                            "          NF.EndEmpresa_Id," & vbCrLf &
                            "          E.Nome AS EmpresaNome," & vbCrLf &
                            "          E.Cidade AS EmpresaCidade," & vbCrLf &
                            "          E.Estado AS EmpresaEstado," & vbCrLf &
                            "          NF.EntradaSaida_Id," & vbCrLf &
                            "   	   P.Grupo," & vbCrLf &
                            "          ge.Descricao as DescGrupo," & vbCrLf &
                            "          NFI.Produto_Id AS Produto," & vbCrLf &
                            "          P.Nome AS NomeProduto," & vbCrLf &
                            "          NF.Operacao," & vbCrLf &
                            "          NF.SubOperacao," & vbCrLf &
                            "          op.Classe," & vbCrLf &
                            "          SO.Descricao as DescOperacao," & vbCrLf &
                            "          ISNULL(dbo.getNomeCliente(CLI.Cliente_Id), '') AS NomeCliente," & vbCrLf &
                            "           CLI.Cidade + '/' + CLI.Estado as Cidade," & vbCrLf &
                            "   	   SUM(nfi.QuantidadeFiscal) PesoFiscal," & vbCrLf &
                            "          SUM(NFI.Valor) as Valor" & vbCrLf &
                            "     FROM NotasFiscais AS NF WITH (NOLOCK)" & vbCrLf &
                            "    INNER JOIN NotasFiscaisXItens AS NFI WITH (NOLOCK)" & vbCrLf &
                            "       ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
                            "      AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
                            "      AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                            "      AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
                            "      AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
                            "      AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf &
                            "      AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf &
                            "     LEFT Join NFERealizadas NFER WITH (NOLOCK)" & vbCrLf &
                            "       ON NF.Empresa_Id      = NFER.Empresa_Id" & vbCrLf &
                            "      AND NF.EndEmpresa_Id   = NFER.EndEmpresa_Id" & vbCrLf &
                            "      AND NF.Cliente_Id      = NFER.Cliente_Id" & vbCrLf &
                            "      AND NF.EndCliente_Id   = NFER.EndCliente_Id" & vbCrLf &
                            "      AND NF.EntradaSaida_Id = NFER.EntradaSaida_Id" & vbCrLf &
                            "      AND NF.Serie_Id        = NFER.Serie_Id" & vbCrLf &
                            "      AND NF.Nota_Id         = NFER.Nota_Id" & vbCrLf &
                            "     LEFT Join Pedidos Pe WITH (NOLOCK)" & vbCrLf &
                            "       on NF.Empresa_Id    = Pe.Empresa_Id" & vbCrLf &
                            "      and NF.EndEmpresa_Id = Pe.EndEmpresa_Id" & vbCrLf &
                            "      and NF.Pedido        = Pe.Pedido_Id" & vbCrLf &
                            "    INNER JOIN Produtos AS P WITH (NOLOCK)" & vbCrLf &
                            "       ON P.Produto_Id = NFI.Produto_Id" & vbCrLf &
                            "    INNER JOIN GruposDeEstoques AS ge WITH (NOLOCK)" & vbCrLf &
                            "       ON ge.Grupo_Id = p.Grupo" & vbCrLf &
                            "     LEFT JOIN Marca M WITH (NOLOCK)" & vbCrLf &
                            "       ON M.Marca_id = P.Marca" & vbCrLf &
                            "    INNER JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id," & vbCrLf &
                            "                       sum(case when Encargo_Id ='LIQUIDO'         THEN isnull(Valor, 0)                      else 0 end) AS LIQUIDO" & vbCrLf &
                            "                  FROM NotasFiscaisXEncargos WITH (NOLOCK)" & vbCrLf &
                            "                 Where Encargo_Id in ('LIQUIDO', 'PRODUTO','ICMS','FUNRURAL','FUNRURAL TERCEI','FETHAB','FACS')" & vbCrLf &
                            "                 group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id,  Sequencia_Id" & vbCrLf &
                            "                ) AS Encargos" & vbCrLf &
                            "       ON NFI.Empresa_Id      = Encargos.Empresa_Id" & vbCrLf &
                            "      AND NFI.EndEmpresa_Id   = Encargos.EndEmpresa_Id" & vbCrLf &
                            "      AND NFI.Cliente_Id      = Encargos.Cliente_Id" & vbCrLf &
                            "      AND NFI.EndCliente_Id   = Encargos.EndCliente_Id" & vbCrLf &
                            "      AND NFI.EntradaSaida_Id = Encargos.EntradaSaida_Id" & vbCrLf &
                            "      AND NFI.Serie_Id        = Encargos.Serie_Id" & vbCrLf &
                            "      AND NFI.Nota_Id         = Encargos.Nota_Id" & vbCrLf &
                            "      AND NFI.Produto_Id      = Encargos.Produto_Id" & vbCrLf &
                            "      AND NFI.CFOP_Id         = Encargos.CFOP_Id" & vbCrLf &
                            "      AND NFI.Sequencia_Id    = Encargos.Sequencia_Id" & vbCrLf &
                            "    Inner Join Operacoes as OP WITH (NOLOCK)" & vbCrLf &
                            "       ON OP.Operacao_Id     = NFI.Operacao" & vbCrLf &
                            "    INNER JOIN SubOperacoes AS SO WITH (NOLOCK)" & vbCrLf &
                            "       ON SO.Operacao_Id     = NFI.Operacao" & vbCrLf &
                            "      AND SO.SubOperacoes_Id = NFI.SubOperacao" & vbCrLf &
                            "     LEFT JOIN NotasFiscaisXRomaneios AS NFR WITH (NOLOCK)" & vbCrLf &
                            "       ON NFR.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
                            "      AND NFR.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
                            "      AND NFR.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
                            "      AND NFR.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
                            "      AND NFR.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
                            "      AND NFR.Serie_Id        = NFI.Serie_Id" & vbCrLf &
                            "      AND NFR.Nota_Id         = NFI.Nota_Id" & vbCrLf &
                            "     LEFT JOIN Clientes AS CLI WITH (NOLOCK)" & vbCrLf &
                            "       ON nf.Cliente_Id    = CLI.Cliente_Id" & vbCrLf &
                            "      AND nf.EndCliente_Id = CLI.Endereco_Id" & vbCrLf &
                            "     LEFT JOIN Clientes AS E WITH (NOLOCK)" & vbCrLf &
                            "       ON nf.Empresa_Id    = E.Cliente_Id" & vbCrLf &
                            "      AND nf.EndEmpresa_Id = E.Endereco_Id" & vbCrLf &
                            "     LEFT JOIN Clientes AS DO WITH (NOLOCK)" & vbCrLf &
                            "       ON NFI.Deposito    = DO.Cliente_Id" & vbCrLf &
                            "      AND NFI.EndDeposito = DO.Endereco_Id" & vbCrLf &
                            "     LEFT JOIN Clientes AS DD WITH (NOLOCK)" & vbCrLf &
                            "       ON NF.Destino    = DD.Cliente_Id" & vbCrLf &
                            "      AND NF.EndDestino = DD.Endereco_Id" & vbCrLf &
                            "    WHERE NF.Situacao = 1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            If Not chkConsolidarEmpresa.Checked Then
                sql &= "    AND NF.Empresa_Id     = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                       "    AND NF.EndEmpresa_Id  =  " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            Else
                sql &= "    AND left(NF.Empresa_Id, 8) ='" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            If Not chkConsolidarEmpresa.Checked Then
                sql &= "    AND NF.Cliente_Id    ='" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                       "    AND NF.EndCliente_Id  = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            Else
                sql &= "    AND left(NF.Cliente_Id, 8) ='" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'" & vbCrLf
            End If
        End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= "And " & ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "NFI.Produto_id")(0) & vbCrLf
        End If

        If ddlSafra.SelectedIndex > 0 Then sql &= "AND PE.Safra = '" & ddlSafra.SelectedValue & "'" & vbCrLf

        If validaPeriodo() Then sql &= "   And NF.Movimento between '" & CDate(txtData1.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtData2.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf

        sql &= "    AND SO.Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf &
               "    AND OP.Classe = '" & ddlClasse.SelectedValue & "'" & vbCrLf &
               "  Group by E.Reduzido, NF.Empresa_Id, NF.EndEmpresa_Id, E.Nome, E.Cidade, E.Estado," & vbCrLf &
               "           NF.EntradaSaida_Id, P.Grupo, ge.Descricao, NFI.Produto_Id," & vbCrLf &
               "           P.Nome, NF.Operacao, NF.SubOperacao, op.Classe, SO.Descricao," & vbCrLf &
               "           ISNULL(dbo.getNomeCliente(CLI.Cliente_Id), ''), CLI.Cidade + '/' + CLI.Estado" & vbCrLf &
               "  order by EmpresaNome, EmpresaCidade, EmpresaEstado, Classe, DescGrupo, NomeCliente, Cidade, " & vbCrLf &
               "           EntradaSaida_Id, Operacao, SubOperacao, DescOperacao, NomeProduto"

        Return Banco.ConsultaDataSet(sql, "MovimentoFiscalPorOperacao")
    End Function

    Private Sub getParametrosConsulta(ByRef param As Dictionary(Of String, Object))
        param.Add("ParametroConsulta", "Unidade: " & ddlUnidade.SelectedItem.Text & vbCrLf)

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim obj As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param("ParametroConsulta") &= "Empresa: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " - " & obj.CodigoEstado & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            Dim obj As New Cliente(txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1))
            param("ParametroConsulta") &= "Cliente: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " - " & obj.CodigoEstado & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlClasse.SelectedValue) Then param("ParametroConsulta") &= "Classe: " & ddlClasse.SelectedValue & vbCrLf

        If ucSelecaoProduto.TemSelecionado Then param("ParametroConsulta") &= ucSelecaoProduto.GetSqlEParametrosRelatorio("Grupo_Id", "Produto_Id", "", True)(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then param("ParametroConsulta") &= "Safra: " & ddlClasse.SelectedValue & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtData1.Text) Then param("ParametroConsulta") &= "Período: " & txtData1.Text & " á " & txtData2.Text

    End Sub

    Private Function validaPeriodo() As Boolean
        Return Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not String.IsNullOrWhiteSpace(txtData2.Text) AndAlso IsDate(txtData1.Text) AndAlso IsDate(txtData2.Text)
    End Function

#End Region

End Class