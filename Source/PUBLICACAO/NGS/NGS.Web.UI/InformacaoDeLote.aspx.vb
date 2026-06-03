Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Drawing

Public Class InformacaoDeLote
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("InformacaoDeLote", "ACESSAR") Then
                CarregarUnidade()
                ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                Limpar()
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
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

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteInfLT" & HID.Value) Is Nothing Then
            txtCodigoCliente.Value = CType(Session("objClienteInfLT" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objClienteInfLT" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(CType(Session("objClienteInfLT" & HID.Value), [Lib].Negocio.Cliente).Codigo) & " - " & CType(Session("objClienteInfLT" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteInfLT" & HID.Value)
        End If
    End Sub

    Function Validar() As Boolean
        If ddlUnidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Unidade de Negocio")
            Return False
        ElseIf DdlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Empresa")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
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

    Private Function VerificarGrupoProduto(ByRef Sql As String) As String
        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "NFI.Produto_id")
            Sql &= " AND " & retorno(0)
            Return retorno(1)
        Else
            Return ""
        End If
    End Function

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio()
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
            If Funcoes.VerificaPermissao("InformacaoDeLote", "RELATORIO") Then
                Dim Periodo As String = "PERÍODO: " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text))
                Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

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
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("INFORMAÇÕES DE LOTE")

                        'criando linha com o cabeçalho da planilha
                        Dim rowIndex As Integer = 1
                        Dim columnIndex As Integer = 1

                        'criando linha que informa o nome da empresa e o cnpj
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa a cidade e o estado da empresa
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o título do relatório
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "INFORMAÇÕES DE LOTE")
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:P" & rowIndex).AutoFilter = True

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using
                        rowIndex += 1

                        ' criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                columnIndex += 1
                            Next

                            'formatando células datas
                            worksheet.Cells(String.Format("A{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                            worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                            worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                            'formatando células numéricas
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"
                            worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000000000_ ;[Red]-#,##0.0000000000"
                            'worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'aplicando formatação nas células do conteúdo
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Formula = String.Format("=SUM(O6:O{0})", rowIndex - 1)
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000_ ;[Red]-#,##0.0000"

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Formula = String.Format("=SUM(P6:P{0})", rowIndex - 1)
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

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

        Sql = " SELECT N.Movimento, N.DataDaNota, N.Nota_Id as Nota, N.Serie_Id as Serie, N.Pedido as Pedido, C.Cliente_Id as Cliente, C.Nome, " & vbCrLf
        Sql &= "		case" & vbCrLf
        Sql &= "			when C.Estado = 'EX'" & vbCrLf
        Sql &= "				then C.Cidade + '-' + p.Descricao" & vbCrLf
        Sql &= "				else C.Cidade + '-' + C.Estado" & vbCrLf
        Sql &= "		end as Origem, " & vbCrLf
        Sql &= "       NFI.Produto_Id as Produto, prd.Nome AS NomeDoProduto, " & vbCrLf
        Sql &= "       NXL.Lote_Id as Lote, NXL.Fabricado, NXL.Validade, NXL.Quantidade, NFI.Unitario, (NXL.Quantidade * NFI.Unitario) as Valor" & vbCrLf
        Sql &= "  FROM  NotasFiscais N" & vbCrLf
        Sql &= "		 INNER JOIN Clientes C" & vbCrLf
        Sql &= "				 ON C.Cliente_Id        = N.Cliente_Id" & vbCrLf
        Sql &= "				AND C.Endereco_Id       = N.EndCliente_Id" & vbCrLf
        Sql &= "		 INNER JOIN Pais p" & vbCrLf
        Sql &= "				 ON P.Pais_Id           = C.Pais" & vbCrLf
        Sql &= "		 INNER JOIN NotasFiscaisXItens NFI" & vbCrLf
        Sql &= "				 ON NFI.Empresa_Id      = N.Empresa_Id" & vbCrLf
        Sql &= "				AND NFI.EndEmpresa_Id   = N.EndEmpresa_Id" & vbCrLf
        Sql &= "				AND NFI.Cliente_Id      = N.Cliente_Id" & vbCrLf
        Sql &= "				AND NFI.EndCliente_Id   = N.EndCliente_Id " & vbCrLf
        Sql &= "				AND NFI.EntradaSaida_Id = N.EntradaSaida_Id" & vbCrLf
        Sql &= "				AND NFI.Serie_Id        = N.Serie_Id" & vbCrLf
        Sql &= "				AND NFI.Nota_Id         = N.Nota_Id" & vbCrLf
        Sql &= "		 INNER JOIN NotaFiscalXLote NXL" & vbCrLf
        Sql &= "				 ON NXL.Empresa_Id       = NFI.Empresa_Id" & vbCrLf
        Sql &= "				 AND NXL.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf
        Sql &= "				 AND NXL.Cliente_Id      = NFI.Cliente_Id" & vbCrLf
        Sql &= "				 AND NXL.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf
        Sql &= "				 AND NXL.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf
        Sql &= "				 AND NXL.Serie_Id        = NFI.Serie_Id" & vbCrLf
        Sql &= "				 AND NXL.Nota_Id         = NFI.Nota_Id" & vbCrLf
        Sql &= "				 AND NXL.Produto_Id      = NFI.Produto_Id" & vbCrLf
        Sql &= "				 AND NXL.CFOP_Id         = NFI.CFOP_Id" & vbCrLf
        Sql &= "				 AND NXL.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf
        Sql &= "		 INNER JOIN Produtos prd" & vbCrLf
        Sql &= "				 ON prd.Produto_Id         = NFI.Produto_Id" & vbCrLf
        Sql &= "        Where N.Situacao = 1 And N.Empresa_Id = '" & objEmpresa.Codigo & "'" & vbCrLf

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= "          And N.Cliente_Id = '" & strCliente(0) & "' And N.EndCliente_Id = " & strCliente(1) & vbCrLf
        End If

        Sql &= "          And N.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf

        VerificarGrupoProduto(Sql)

        If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND N.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf
        If cmbSubOperacao.SelectedIndex > 0 Then
            Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            Sql &= "AND N.SubOperacoes = " & strSubOpe(1) & " " & vbCrLf
        End If

        If rdValidade.Checked Then
            Sql &= "        Order By NXL.Validade" & vbCrLf
        ElseIf rdMovimento.Checked Then
            Sql &= "        Order By N.Movimento" & vbCrLf
        ElseIf rdProduto.Checked Then
            Sql &= "        Order By NFI.Produto_Id" & vbCrLf
        Else
            Sql &= "        Order By C.Nome" & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "InformacaoDeLote")

        Return ds
    End Function

End Class