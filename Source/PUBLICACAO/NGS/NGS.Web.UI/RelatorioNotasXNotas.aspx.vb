Imports System.Drawing
Imports System.IO
Imports NGS.Lib.Negocio
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class RelatorioNotasXNotas
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioNotasXNotas", "ACESSAR") Then
                CarregarUnidade()
                ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                Limpar()
                LiberaEmpresa()
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                CarregarTipoDeDocumento(False)
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

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteCDNF" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
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

    Private Sub CarregarTipoDeDocumento(ByVal todos As Boolean)
        chkTipoDeDocumento.Items.Clear()
        chkTipoDeDocumento.DataValueField = "Codigo"
        chkTipoDeDocumento.DataTextField = "Descricao"
        Dim lst As New [Lib].Negocio.ListTipoDeDocumento()
        chkTipoDeDocumento.DataSource = lst.ToArray()
        chkTipoDeDocumento.DataBind()
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If todos Then
                chkTipoDeDocumento.Items(i).Selected = True
            ElseIf chkTipoDeDocumento.Items(i).Value = 1 Then
                chkTipoDeDocumento.Items(i).Selected = True
            End If
        Next
    End Sub

    Protected Sub chkAllTipos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkAllTipos.Checked Then
                For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
                    chkTipoDeDocumento.Items(i).Selected = True
                Next
            Else
                For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1

                    If chkTipoDeDocumento.Items(i).Value = 1 Then
                        chkTipoDeDocumento.Items(i).Selected = True
                    Else
                        chkTipoDeDocumento.Items(i).Selected = False
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteCDNF" & HID.Value) Is Nothing Then
            txtCodigoCliente.Value = CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).Codigo) & " - " & CType(Session("objClienteCDNF" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteCDNF" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
    End Sub

    Protected Sub lnkRelatório_Click(sender As Object, e As EventArgs) Handles lnkRelatório.Click
        Try
            EmitirRelatorio()
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
            Funcoes.Ajuda(Me.Page, "ConsistenciaDeNotas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub EmitirRelatorio()
        Try
            If Funcoes.VerificaPermissao("RelatorioNotasXNotas", "RELATORIO") Then
                Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))
                Dim data As String = CStr(txtDataInicial.Text & " - " & txtDataFinal.Text)
                Dim ds As DataSet = getDataSet(objEmpresa)
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
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório De Notas X Notas.")

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
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Notas X Notas")
                        worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & data)
                        worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o subtitulo Origem/Destino
                        'criando o subtitulo Origem
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0))
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Value = String.Format("{0}", "Origem")
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:S{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                        'criando o subtitulo Destino
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 224, 180))
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Value = String.Format("{0}", "Destino")
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("T{0}:AL{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        rowIndex += 1

                        ' criando cabeçalho da planilha com os dados do dataset
                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In dt.Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A6:AD" & rowIndex).AutoFilter = True

                        'formatando células numéricas
                        'QuantidadeFiscal
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#.##0.0000_ ;[Red]-#.##0.0000"
                        'Unitário
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#.##0.0000000000_ ;[Red]-#.##0.0000000000"
                        'Valor
                        worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "#.##0.00_ ;[Red]-#.##0.00"
                        'Quantidade
                        worksheet.Cells(String.Format("AH{0}", rowIndex)).Style.Numberformat.Format = "#.##0.0000_ ;[Red]-#.##0.0000"
                        'Unitário1
                        worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Numberformat.Format = "#.##0.0000000000_ ;[Red]-#.##0.0000000000"
                        'Valor1
                        worksheet.Cells(String.Format("AJ{0}", rowIndex)).Style.Numberformat.Format = "#.##0.00_ ;[Red]-#.##0.00"

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
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir o relatorio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataSet(ByVal objEmpresa As [Lib].Negocio.Cliente) As DataSet

        Sql = "SELECT n.Empresa_Id, n.EndEmpresa_Id, n.Cliente_Id, n.EndCliente_Id, CNFO.Nome, n.EntradaSaida_Id, n.Serie_Id, n.Nota_Id, " & vbCrLf &
        "       n.Movimento, n.Pedido, n.Operacao, n.SubOperacao, SBO.Descricao, pO.Nome, NFI.QuantidadeFiscal, NFI.Unitario, NFI.Valor, n.TipoDeDocumento, TdNFO.Descricao, " & vbCrLf &
        "       nDES.Empresa_Id, nDES.EndEmpresa_Id, nDES.Cliente_Id, nDES.EndCliente_Id, CNFD.Nome, nDES.EntradaSaida_Id, nDES.Serie_Id, nDES.Nota_Id, " & vbCrLf &
        "	   nDES.Movimento, nDES.Pedido, nDES.Operacao, nDES.SubOperacao, SBO.Descricao, pD.Nome, isnull(NFIDES.QuantidadeFiscal,0) AS Quantidade, " & vbCrLf &
        "      isnull(NFIDES.Unitario,0) AS Unitario, isnull(NFIDES.Valor,0) AS Valor, nDES.TipoDeDocumento, TdNFD.Descricao " & vbCrLf &
        "FROM NotasFiscais n " & vbCrLf &
        "	INNER JOIN NotasFiscaisXItens AS NFI " & vbCrLf &
        "			 ON NFI.Empresa_Id      = n.Empresa_Id " & vbCrLf &
        "			AND NFI.EndEmpresa_Id   = n.EndEmpresa_Id " & vbCrLf &
        "			AND NFI.Cliente_Id      = n.Cliente_Id " & vbCrLf &
        "			AND NFI.EndCliente_Id   = n.EndCliente_Id " & vbCrLf &
        "			AND NFI.EntradaSaida_Id = n.EntradaSaida_Id " & vbCrLf &
        "			AND NFI.Serie_Id        = n.Serie_Id " & vbCrLf &
        "			AND NFI.Nota_Id         = n.Nota_Id " & vbCrLf &
        "	INNER JOIN Produtos pO " & vbCrLf &
        "			 ON pO.Produto_Id  = NFI.Produto_Id " & vbCrLf &
        "	INNER JOIN TipoDeDocumento TdNFO " & vbCrLf &
        "			 ON TdNFO.Codigo_Id  = n.TipoDeDocumento  " & vbCrLf &
        "	INNER JOIN Clientes CNFO " & vbCrLf &
        "			 ON CNFO.Cliente_Id  = n.Cliente_Id  " & vbCrLf &
        "			AND CNFO.Endereco_Id = n.EndCliente_Id  " & vbCrLf &
        "	INNER JOIN SubOperacoes SBO " & vbCrLf &
        "			 ON SBO.Operacao_id     = n.Operacao  " & vbCrLf &
        "			AND SBO.SubOperacoes_id = n.SubOperacao  " & vbCrLf &
        "	left join NotasXNotas AS NxN " & vbCrLf &
        "			 ON NxN.OrigemEmpresa_Id      = n.Empresa_Id " & vbCrLf &
        "			AND NxN.OrigemEndEmpresa_Id   = n.EndEmpresa_Id " & vbCrLf &
        "			AND NxN.OrigemCliente_Id      = n.Cliente_Id " & vbCrLf &
        "			AND NxN.OrigemEndCliente_Id   = n.EndCliente_Id " & vbCrLf &
        "			AND NxN.OrigemEntradaSaida_Id = n.EntradaSaida_Id " & vbCrLf &
        "			AND NxN.OrigemSerie_Id        = n.Serie_Id " & vbCrLf &
        "			AND NxN.OrigemNota_Id         = n.Nota_Id " & vbCrLf &
        "	left join NotasFiscais AS nDES " & vbCrLf &
        "			 ON nDES.Empresa_Id      = NxN.Empresa_Id " & vbCrLf &
        "			AND nDES.EndEmpresa_Id   = NxN.EndEmpresa_Id " & vbCrLf &
        "			AND nDES.Cliente_Id      = NxN.Cliente_Id " & vbCrLf &
        "			AND nDES.EndCliente_Id   = NxN.EndCliente_Id " & vbCrLf &
        "			AND nDES.EntradaSaida_Id = NxN.EntradaSaida_Id " & vbCrLf &
        "			AND nDES.Serie_Id        = NxN.Serie_Id " & vbCrLf &
        "			AND nDES.Nota_Id         = NxN.Nota_Id " & vbCrLf &
        "	LEFT JOIN NotasFiscaisXItens AS NFIDES " & vbCrLf &
        "			 ON NFIDES.Empresa_Id      = nDES.Empresa_Id " & vbCrLf &
        "			AND NFIDES.EndEmpresa_Id   = nDES.EndEmpresa_Id " & vbCrLf &
        "			AND NFIDES.Cliente_Id      = nDES.Cliente_Id " & vbCrLf &
        "			AND NFIDES.EndCliente_Id   = nDES.EndCliente_Id " & vbCrLf &
        "			AND NFIDES.EntradaSaida_Id = nDES.EntradaSaida_Id " & vbCrLf &
        "			AND NFIDES.Serie_Id        = nDES.Serie_Id " & vbCrLf &
        "			AND NFIDES.Nota_Id         = nDES.Nota_Id " & vbCrLf &
        "	LEFT JOIN TipoDeDocumento TdNFD " & vbCrLf &
        "			 ON TdNFD.Codigo_Id     = nDES.TipoDeDocumento  " & vbCrLf &
        "	LEFT JOIN Produtos pD " & vbCrLf &
        "			 ON pD.Produto_Id  = NFIDES.Produto_Id " & vbCrLf &
        "	LEFT JOIN Clientes CNFD " & vbCrLf &
        "			 ON CNFD.Cliente_Id  = nDES.Cliente_Id  " & vbCrLf &
        "			AND CNFD.Endereco_Id = nDES.EndCliente_Id  " & vbCrLf &
        "	LEFT JOIN SubOperacoes SBD " & vbCrLf &
        "			 ON SBD.Operacao_id     = nDES.Operacao  " & vbCrLf &
        "			AND SBD.SubOperacoes_id = nDES.SubOperacao  " & vbCrLf &
        "Where n.Situacao = 1 " & vbCrLf

        If chkUnificarEmpresa.Checked Then
            Sql &= "        And n.Empresa_Id = '" & Left(objEmpresa.Codigo, 8) & "'" & vbCrLf
        Else
            Sql &= "        And n.Empresa_Id    = '" & objEmpresa.Codigo & "'" & vbCrLf
            Sql &= "        And n.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & vbCrLf
        End If

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= "          And n.Cliente_Id = '" & strCliente(0) & "' And n.EndCliente_Id = " & strCliente(1) & vbCrLf
        End If

        Sql &= "          And n.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
        Sql &= "          And n.EntradaSaida_Id = '" & IIf(RadEntradas.Checked, "E", "S") & "'" & vbCrLf

        VerificarGrupoProduto(Sql)

        If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND N.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf
        If cmbSubOperacao.SelectedIndex > 0 Then
            Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            Sql &= "AND N.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
        End If

        Dim tp As String = ""
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If chkTipoDeDocumento.Items(i).Selected Then
                tp &= IIf(tp.Length = 0, "", ",") & chkTipoDeDocumento.Items(i).Value
            End If
        Next

        If tp <> "" Then
            Sql &= " AND n.TipoDeDocumento in(" & tp & ")" & vbCrLf
        End If

        If chkNossaEmissao.Checked Then
            Sql &= "AND n.NossaEmissao = 'S'" & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "RelatorioNotasXNotas")
        Return ds
    End Function

End Class