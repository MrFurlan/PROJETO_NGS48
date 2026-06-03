Imports System.IO
Imports System.Drawing
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioCtrcDasNotas
    Inherits BasePage

#Region "Atributos"
    Private objCliente As [Lib].Negocio.Cliente
#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioCtrcDasNotas", "ACESSAR") Then
                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioCtrcDasNotas", "RELATORIO") Then
                If IsValidFields() Then

                    Dim ds As DataSet = getDataSet()
                    Dim parametros As String = getParametros()
                    Dim param As New Dictionary(Of String, Object)()
                    param.Add("Titulo", "Relatório de CTRC das Notas.")
                    param.Add("parametros", parametros)
                    param.Add("cab1", IIf(rdoNotaXCtrc.Checked, "NOTA FISCAL", "CTRC CIRCULAÇÃO"))
                    param.Add("cab2", IIf(rdoNotaXCtrc.Checked, "CONHECIMENTO DE TRANSPORTE", "CTRC COMPROVAÇÃO"))

                    Funcoes.BindReport(Me.Page, ds, "Cr_CtrcNotas", eExportType.PDF, param)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permnissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaCliente_Click(sender As Object, e As EventArgs) Handles btnConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCN" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaTransp_Click(sender As Object, e As EventArgs) Handles btnConsultaTransp.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
            Popup.ConsultaDeClientes(Me.Page, "objClienteTransp" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(sender As Object, e As EventArgs) Handles btnPedido.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada!")
            ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "Cliente não foi selecionado!")
            Else
                HttpContext.Current.Session("ssCampo") = "Pedidos"
                Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                Dim parameters As New Dictionary(Of String, Object)
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                    parameters("pedido") = txtPedido.Text.Trim()
                End If
                Popup.ConsultaDePedidos(Me.Page, "objCtrcDasNotas" & HID.Value, "txtNome")
                ucConsultaPedidos.BindGridView(parameters)
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

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioCtrcDasNotas", "RELATORIO") Then
                If IsValidFields() Then
                    Dim ds As New DataSet
                    ds = getDataSet()

                    If ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                        If File.Exists(fileName) Then
                            File.Delete(fileName)
                        End If

                        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                            Using package As New ExcelPackage(arquivo)

                                Dim worksheet As ExcelWorksheet
                                Dim rowIndex As Integer = 1
                                Dim columnIndex As Integer = 1

                                If rdoNotaXCtrc.Checked Then
                                    worksheet = package.Workbook.Worksheets.Add("NOTAFISCAL - CTRC")
                                Else
                                    worksheet = package.Workbook.Worksheets.Add("CTRC CIRC./CTRC COMPROV.")
                                End If

                                'Adiciona Colunas
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName.Replace("_Id", "").Replace("EndEmpresa", "End").Replace("EndCliente", "End").Replace("EntradaSaida", "E/S")
                                    columnIndex += 1
                                Next

                                'aplicando formatação nas células das Colunas
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Font.Bold = True
                                    range.Style.Font.Size = 12
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                    range.Style.Font.Color.SetColor(Color.White)
                                End Using

                                rowIndex += 1

                                For Each row As DataRow In ds.Tables(0).Rows
                                    columnIndex = 1
                                    For Each col As DataColumn In ds.Tables(0).Columns
                                        If col.ColumnName.Contains("Movimento") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "mm/dd/yyyy"
                                        ElseIf col.ColumnName = "Valor" Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Border.Right.Style = ExcelBorderStyle.Thin
                                            worksheet.Cells(rowIndex, columnIndex).Style.Border.Right.Color.SetColor(Color.FromArgb(0, 0, 0))
                                        ElseIf col.ColumnName.Contains("End") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                        End If

                                        If col.ColumnName.Contains("Valor") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                        ElseIf col.ColumnName.Contains("Serie") OrElse col.ColumnName.Contains("Quantidade") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                        End If

                                        If col.ColumnName.Contains("Serie") Then
                                            If IsNumeric(row(col.ColumnName)) Then
                                                worksheet.Cells(rowIndex, columnIndex).Value = CInt(row(col.ColumnName))
                                            Else
                                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                            End If

                                        Else
                                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                        End If

                                        columnIndex += 1

                                    Next

                                    If rowIndex Mod 2 = 0 Then
                                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                        End Using
                                    End If

                                    rowIndex += 1
                                Next

                                'Aplicando formatação no rodapé
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Font.Bold = True
                                    range.Style.Font.Size = 12
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                    range.Style.Font.Color.SetColor(Color.White)
                                End Using

                                'criando auto filtro na planilha
                                worksheet.Cells("A1:X" & rowIndex).AutoFilter = True

                                'setando autofit nas células da planilha
                                worksheet.Cells.AutoFitColumns(0)

                                'congelando primeira linha
                                worksheet.View.FreezePanes(2, 1)

                                'salvando planilha do excel
                                package.Save()
                            End Using
                        End Using
                        'download do arquivo pelo browser
                        Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                    Else
                        MsgBox(Me.Page, "Nenhum resultado encontrado, com os parâmetros informado.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Function IsValidFields() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            ddlEmpresa.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData1.Text) OrElse String.IsNullOrWhiteSpace(txtData2.Text) Then
            MsgBox(Me.Page, "Informe as datas para consulta.")
            Return False
        End If
        Return True
    End Function

    Private Function getDataSet() As DataSet
        Dim ds As New DataSet
        Dim sql As String = ""

        sql &= "   SELECT  " & vbCrLf &
            "   	nf.Pedido," & vbCrLf &
            "   	nf.Nota_Id," & vbCrLf &
            "   	nf.Serie_Id," & vbCrLf &
            "   	nf.EntradaSaida_Id," & vbCrLf &
            "   	nxi.Produto_Id," & vbCrLf &
            "   	Prd.Nome As ProdutoNomeNota," & vbCrLf &
            "   	CONCAT(sub.Operacao_Id, '-', sub.SubOperacoes_Id,  '-',sub.Descricao) AS OperacaoNota, " & vbCrLf &
            "   	nf.Empresa_Id As Empresa_Id," & vbCrLf &
            "   	nf.EndEmpresa_Id," & vbCrLf &
            "   	nf.Cliente_Id As Cliente_Id," & vbCrLf &
            "   	nf.EndCliente_Id," & vbCrLf &
            "   	cli1.Nome As Cliente," & vbCrLf &
            "   	nf.Movimento," & vbCrLf &
            "   	nxi.QuantidadeFiscal As Quantidade," & vbCrLf &
            "   	nxi.Valor," & vbCrLf &
            "   	cte.Nota_Id As Ctrc_Nota_Id," & vbCrLf &
            "   	cte.Serie_Id As Ctrc_Serie_Id," & vbCrLf &
            "   	cte.EntradaSaida_Id As Ctrc_EntradaSaida_Id," & vbCrLf &
            "   	ctexi.Produto_Id As ProdutoCTE," & vbCrLf &
            "   	Prd_CTE.Nome As ProdutoNomeCTE," & vbCrLf &
            "   	CONCAT(sub_CTE.Operacao_Id, '-', sub_CTE.SubOperacoes_Id,  '-',sub_CTE.Descricao) AS OperacaoCTE, " & vbCrLf &
            "   	cte.Empresa_Id As Ctrc_Empresa_Id," & vbCrLf &
            "   	cte.EndEmpresa_Id As Ctrc_EndEmpresa_Id," & vbCrLf &
            "   	cte.Cliente_Id As Ctrc_Cliente_Id," & vbCrLf &
            "   	cte.EndCliente_Id As Ctrc_EndCliente_Id," & vbCrLf &
            "   	cli2.Nome As Ctrc_Cliente," & vbCrLf &
            "   	cte.Movimento As Ctrc_Movimento," & vbCrLf &
            "   	ctexi.QuantidadeFiscal As Ctrc_Quantidade," & vbCrLf &
            "   	ctexi.Unitario As Ctrc_Unitario," & vbCrLf &
            "   	ctexi.Valor As Ctrc_Valor," & vbCrLf &
            "   	nxt.Proprietario + ' - ' + trp.Nome as Transportador_Id" & vbCrLf &
            "   FROM NotasFiscais AS nf " & vbCrLf &
            "  INNER JOIN NotasFiscaisXItens AS nxi" & vbCrLf &
            "     ON nf.Empresa_Id = nxi.Empresa_Id" & vbCrLf &
            "    AND nf.EndEmpresa_Id = nxi.EndEmpresa_Id" & vbCrLf &
            "    AND nf.Cliente_Id = nxi.Cliente_Id" & vbCrLf &
            "    AND nf.EndCliente_Id = nxi.EndCliente_Id" & vbCrLf &
            "    AND nf.EntradaSaida_Id = nxi.EntradaSaida_Id" & vbCrLf &
            "    AND nf.Serie_Id = nxi.Serie_Id" & vbCrLf &
            "    AND nf.Nota_Id = nxi.Nota_Id" & vbCrLf &
            "  INNER JOIN Clientes cli1" & vbCrLf &
            "     ON cli1.Cliente_Id = nf.Cliente_Id" & vbCrLf &
            "    AND cli1.Endereco_Id = nf.EndCliente_Id" & vbCrLf &
            "  INNER JOIN SubOperacoes sub" & vbCrLf &
            "     ON nxi.Operacao = sub.Operacao_Id" & vbCrLf &
            "    AND nxi.SubOperacao = sub.SubOperacoes_Id" & vbCrLf &
            "  INNER JOIN Produtos Prd " & vbCrLf &
            "    On Prd.Produto_Id = nxi.Produto_Id " & vbCrLf &
            "  INNER JOIN NotasFiscaisXTransportadores nxt" & vbCrLf &
            "     On nf.Empresa_Id = nxt.Empresa_Id" & vbCrLf &
            "    And nf.EndEmpresa_Id = nxt.EndEmpresa_Id" & vbCrLf &
            "    And nf.Cliente_Id = nxt.Cliente_Id" & vbCrLf &
            "    And nf.EndCliente_Id = nxt.EndCliente_Id" & vbCrLf &
            "    And nf.EntradaSaida_Id = nxt.EntradaSaida_Id" & vbCrLf &
            "    And nf.Serie_Id = nxt.Serie_Id" & vbCrLf &
            "    And nf.Nota_Id = nxt.Nota_Id" & vbCrLf &
            "  INNER JOIN Clientes trp" & vbCrLf &
            "     On trp.Cliente_Id = nxt.Proprietario" & vbCrLf &
            "    And trp.Endereco_Id = nxt.EndProprietario" & vbCrLf &
            "   LEFT JOIN NotasXNotas nxn" & vbCrLf &
            "     On nxn.OrigemEmpresa_Id = nf.Empresa_Id" & vbCrLf &
            "    And nxn.OrigemEndEmpresa_Id = nf.EndEmpresa_Id" & vbCrLf &
            "    And nxn.OrigemCliente_Id = nf.Cliente_Id" & vbCrLf &
            "    And nxn.OrigemEndCliente_Id = nf.EndCliente_Id" & vbCrLf &
            "    And nxn.OrigemEntradaSaida_Id = nf.EntradaSaida_Id" & vbCrLf &
            "    And nxn.OrigemSerie_Id = nf.Serie_Id" & vbCrLf &
            "    And nxn.OrigemNota_Id = nf.Nota_Id" & vbCrLf &
            "    And nxn.OrigemSerie_Id <> 'REC'" & vbCrLf &
            "    AND nxn.Serie_Id <> 'REC'" & vbCrLf &
            "   LEFT JOIN NotasFiscais cte" & vbCrLf &
            "     ON nxn.Empresa_Id = cte.Empresa_Id" & vbCrLf &
            "    AND nxn.EndEmpresa_Id = cte.EndEmpresa_Id" & vbCrLf &
            "    AND nxn.Cliente_Id = cte.Cliente_Id" & vbCrLf &
            "    AND nxn.EndCliente_Id = cte.EndCliente_Id" & vbCrLf &
            "    AND nxn.EntradaSaida_Id = cte.EntradaSaida_Id" & vbCrLf &
            "    AND nxn.Serie_Id = cte.Serie_Id" & vbCrLf &
            "    AND nxn.Nota_Id = cte.Nota_Id" & vbCrLf &
            "    AND cte.TipoDeDocumento in(2,57)" & vbCrLf &
            "    AND cte.Serie_Id <> 'REC'" & vbCrLf &
            "   LEFT JOIN NotasFiscaisXItens ctexi" & vbCrLf &
            "     ON cte.Empresa_Id = ctexi.Empresa_Id" & vbCrLf &
            "    AND cte.EndEmpresa_Id = ctexi.EndEmpresa_Id" & vbCrLf &
            "    AND cte.Cliente_Id = ctexi.Cliente_Id" & vbCrLf &
            "    AND cte.EndCliente_Id = ctexi.EndCliente_Id" & vbCrLf &
            "    AND cte.EntradaSaida_Id = ctexi.EntradaSaida_Id" & vbCrLf &
            "    AND cte.Serie_Id = ctexi.Serie_Id" & vbCrLf &
            "    AND cte.Nota_Id = ctexi.Nota_Id" & vbCrLf &
            "   LEFT JOIN Clientes cli2" & vbCrLf &
            "     ON cli2.Cliente_Id = cte.Cliente_Id" & vbCrLf &
            "    AND cli2.Endereco_Id = cte.EndCliente_Id" & vbCrLf &
            "   INNER JOIN SubOperacoes sub_CTE" & vbCrLf &
            "     ON cte.Operacao = sub_CTE.Operacao_Id" & vbCrLf &
            "     AND cte.SubOperacao = sub_CTE.SubOperacoes_Id" & vbCrLf &
            "   INNER Join Produtos Prd_CTE " & vbCrLf &
            "     ON Prd_CTE.Produto_Id = ctexi.Produto_Id " & vbCrLf &
            "  WHERE nf.Situacao In (1,4,7)" & vbCrLf &
            "    And nf.TipoDeDocumento = " & IIf(rdoNotaXCtrc.Checked, "1", "2") & " " & vbCrLf &
            "    And nf.EntradaSaida_Id = " & IIf(rdoEntrada.Checked, "'E'", "'S'") & " " & vbCrLf &
            "    AND nf.CIFFOB = " & IIf(rdoEntrada.Checked, "'FOB'", "'CIF'") & " " & vbCrLf &
            "    AND nf.Movimento BETWEEN '" & String.Format("{0:yyyy-MM-dd}", CDate(txtData1.Text)) & "' AND '" & String.Format("{0:yyyy-MM-dd}", CDate(txtData2.Text)) & "'" & vbCrLf &
            "    AND nf.Serie_Id <> 'REC' " & vbCrLf &
            "    AND sub.Classe in ('" & eClassesOperacoes.AFIXAR.ToString & "', '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "', '" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.COMPRAS.ToString & "') " & vbCrLf

        If Not chkUnificarEmpresa.Checked Then
            sql &= "    AND nf.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   "    AND nf.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not rdoTodos.Checked Then
            sql &= "           AND " & IIf(rdoFisica.Checked, "len(nxt.Proprietario) <= 11", "len(nxt.Proprietario) > 11") & " " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            sql &= "    AND nf.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                   "    AND nf.EndCliente_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoTransp.Value) Then
            sql &= "    AND nxt.Proprietario = '" & txtCodigoTransp.Value.Split("-")(0) & "'" & vbCrLf & _
                   "    AND nxt.EndProprietario = " & txtCodigoTransp.Value.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNota.Text) Then
            sql &= "    AND nf.Nota_Id = " & txtNota.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            sql &= "    AND nf.Pedido = " & txtPedido.Text & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= "   AND  " & ucSelecaoProduto.GetSqlEParametrosRelatorio("nxi.Grupo_Id", "nxi.Produto_Id", "", True)(0) & vbCrLf
        End If

        sql &= "   ORDER BY nf.Nota_Id, nf.Serie_Id, nf.Movimento " & vbCrLf
        ds = Banco.ConsultaDataSet(sql, "CtrcNotas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim cnpj As String = Funcoes.FormatarCpfCnpj(row("Transportador_Id").ToString().Split("-")(0))
            Dim nome As String = row("Transportador_Id").ToString().Substring(row("Transportador_Id").ToString().IndexOf("-"), row("Transportador_Id").ToString().Length - row("Transportador_Id").ToString().IndexOf("-"))
            row("Transportador_Id") = String.Format("{0} {1}", cnpj, nome)
            row("Cliente_Id") = Funcoes.FormatarCpfCnpj(row("Cliente_Id"))
            row("Empresa_Id") = Funcoes.FormatarCpfCnpj(row("Empresa_Id"))
            If Not IsDBNull(row("Ctrc_Cliente_Id")) Then
                row("Ctrc_Cliente_Id") = Funcoes.FormatarCpfCnpj(row("Ctrc_Cliente_Id"))
            End If
            If Not IsDBNull(row("Ctrc_Empresa_Id")) Then
                row("Ctrc_Empresa_Id") = Funcoes.FormatarCpfCnpj(row("Ctrc_Empresa_Id"))
            End If
        Next

        Return ds
    End Function

    Private Function getParametros() As String
        Dim parametros As String = String.Empty

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedItem.Text) Then
            parametros = "EMPRESA: " & ddlEmpresa.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            parametros &= "CLIENTE: " & txtCliente.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtTransportador.Text) Then
            parametros &= "TRANSPORTADOR: " & txtTransportador.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtNota.Text) Then
            parametros &= "NOTA: " & txtNota.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            parametros &= " - PEDIDO: " & txtPedido.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtData1.Text) Then
            parametros &= "PERÍODO: " & txtData1.Text & " Á " & txtData2.Text
        End If

        Return parametros
    End Function

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objClienteCN" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteCN" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteCN" & HID.Value)
        ElseIf Not Session("objClienteTransp" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteTransp" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtTransportador.Text = itemCliente.Text
            txtCodigoTransp.Value = itemCliente.Value
            Session.Remove("objClienteTransp" & HID.Value)
        ElseIf Session("objCtrcDasNotas" & HID.Value) IsNot Nothing Then
            Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objCtrcDasNotas" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            Session.Remove("objCtrcDasNotas" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
        Funcoes.VerificaEmpresa(ddlEmpresa)
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtTransportador.Text = String.Empty
        txtCodigoTransp.Value = String.Empty
        txtNota.Text = String.Empty
        txtPedido.Text = String.Empty
        txtData1.Text = String.Format("01/{0:00}/{1:0000}", DateTime.Now.Month, DateTime.Now.Year)
        txtData2.Text = DateTime.Now.ToShortDateString()
        rdoNotaXCtrc.Checked = True
        rdoCircXComp.Checked = False
        rdoEntrada.Checked = True
        rdoSaida.Checked = False
        rdoFisica.Checked = False
        rdoJuridica.Checked = True
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucSelecaoProduto.Limpar()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioCtrcDasNotas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class