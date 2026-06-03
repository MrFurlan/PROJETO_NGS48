Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class RelatorioDeLaudosPorPlacas
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeLaudosPorPlacas", "ACESSAR") Then
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    BuncarUnidadeDeNegocio()
                    LiberaEmpresa()
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objClienteRelPlaCli" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteRelPlaCli" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteRelPlaCli" & HID.Value)

            ElseIf Not Session("objClienteRelPlaTra" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteRelPlaTra" & HID.Value), [Lib].Negocio.Cliente))
                txtTransportador.Text = itemCliente.Text
                txtCodigoTransportador.Value = itemCliente.Value
                Session.Remove("objClienteRelPlaTra" & HID.Value)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa, True)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClienteRelPlaCli" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTransportador.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClienteRelPlaTra" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub GerarExcel(ByVal ds As DataSet, ByVal objEmpresa As [Lib].Negocio.Cliente)
        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
            Using package As New ExcelPackage(arquivo)

                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Títulos")
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1

                'Montar Cabeçalho
                worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Nome
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                rowIndex += 1
                worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Cidade & "/" & objEmpresa.CodigoEstado
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                rowIndex += 1
                worksheet.Cells(rowIndex, columnIndex).Value = "Relatório de Laudos por Placas"
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                rowIndex += 2

                'aplicando formatação nas células das Colunas
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 1)
                    range.Style.Font.Bold = True
                    range.Style.Font.Size = 12
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    range.Style.Font.Color.SetColor(Color.White)
                End Using

                For Each col As DataColumn In ds.Tables(0).Columns
                    If col.ColumnName <> "Row" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    End If
                Next

                'CENTRALIZAR AS COLUNAS
                'worksheet.Cells(rowIndex, 1, rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                rowIndex += 1

                For Each row As DataRow In ds.Tables(0).Rows
                    columnIndex = 1
                    For Each col As DataColumn In ds.Tables(0).Columns
                        If col.ColumnName <> "Row" Then
                            If col.ColumnName = "Entrada Balança" OrElse col.ColumnName = "Saída Balança" Then
                                If IsDate(row(col.ColumnName)) Then worksheet.Cells(rowIndex, columnIndex).Value = CDate(row(col.ColumnName)).ToString("dd/MM/yyyy HH:mm:ss")
                            ElseIf col.ColumnName = "IE Transportadora" Then
                                If String.IsNullOrWhiteSpace(row(col.ColumnName)) Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = 0
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                End If
                                worksheet.Cells("T" & rowIndex).Style.Numberformat.Format = "@"
                                worksheet.Cells("S" & rowIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If

                            If col.ColumnName = "Vencimento Crtc" AndAlso ds.Tables(0).Rows(rowIndex - 6)("Financeiro") = "PAGO" Then worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)

                            columnIndex += 1
                        End If
                    Next

                    If rowIndex Mod 2 = 0 Then
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 1)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                        End Using
                    End If

                    rowIndex += 1
                Next

                worksheet.Cells(String.Format("K:P")).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                worksheet.Cells(String.Format("Q:Q")).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                worksheet.Cells(String.Format("S:S")).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                worksheet.Cells(String.Format("U:U")).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                worksheet.Cells(String.Format("V:V")).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                worksheet.Cells("M" & rowIndex).Formula = "=SUM(M6:M" & rowIndex - 1 & ")"
                worksheet.Cells("N" & rowIndex).Formula = "=SUM(N6:N" & rowIndex - 1 & ")"
                worksheet.Cells("O" & rowIndex).Formula = "=SUM(O6:O" & rowIndex - 1 & ")"
                worksheet.Cells("Q" & rowIndex).Formula = "=SUM(Q6:Q" & rowIndex - 1 & ")"
                worksheet.Cells("V" & rowIndex).Formula = "=SUM(V6:V" & rowIndex - 1 & ")"

                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count - 1)
                    range.Style.Font.Bold = True
                    range.Style.Font.Size = 12
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    range.Style.Font.Color.SetColor(Color.White)
                End Using

                'criando auto filtro na planilha
                worksheet.Cells("A5:X" & rowIndex).AutoFilter = True

                'setando autofit nas células da planilha
                worksheet.Cells.AutoFitColumns(0)

                'congelando primeira linha
                worksheet.View.FreezePanes(6, 1)

                'salvando planilha do excel
                package.Save()
            End Using
        End Using
        'download do arquivo pelo browser
        Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
    End Sub

    Function GetDatasetExcel()
        Dim sql As String = ""

        sql = "SELECT P.Placa," & vbCrLf &
                 "  ROW_NUMBER() OVER(PARTITION BY P.Pesagem_id ORDER BY P.Pesagem_id DESC) AS Row, " & vbCrLf &
                 "   case " & vbCrLf &
                 "	    when isnull(R.Processo,'') = '' " & vbCrLf &
                 "		then convert(varchar,isnull(P.Pesagem_Id,0))  " & vbCrLf &
                 "		else convert(varchar,isnull(P.Pesagem_Id,0)) + '-' + R.Processo    " & vbCrLf &
                 "    end as Pesagem_Id,    " & vbCrLf &
                 "   isnull(CONVERT(VARCHAR,P.MOVIMENTO,103), '')AS Movimento,  " & vbCrLf &
                 "   isnull(CONVERT(VARCHAR,P.ENTRADABALANCA,120), '') AS 'Entrada Balança',           " & vbCrLf &
                 "   isnull(CONVERT(VARCHAR,P.SAIDABALANCA,120),'') AS 'Saída Balança',                " & vbCrLf &
                 "   case when isnull(R.Processo,'') = ''                                              " & vbCrLf &
                 "		then   isnull(P.PEDIDO,0)                                                      " & vbCrLf &
                 "		else isnull(R.PEDIDO,0)                                                        " & vbCrLf &
                 "   end AS Pedido,                                                                    " & vbCrLf &
                 "   P.Empresa_Id as 'CNPJ Empresa',                                                   " & vbCrLf &
                 "   CL.NOME AS 'Nome Cliente',                                                        " & vbCrLf &
                 "   PR.NOME AS 'Nome Produto',                                                        " & vbCrLf &
                 "   P.EntradaSaida as 'E/S',                                                          " & vbCrLf &
                 "   P.PrimeiraPesagem,                                                                " & vbCrLf &
                 "   P.SegundaPesagem,                                                                 " & vbCrLf &
                    "   --end as SegundaPesagem,                                                       " & vbCrLf &
                 "    case                                                                             " & vbCrLf &
                 "	when isnull(R.Processo,'') = ''                                                    " & vbCrLf &
                 "		then isnull(P.BrutoBalanca,0)                                                  " & vbCrLf &
                 "		else isnull(R.PesoBruto,0)                                                     " & vbCrLf &
                 "    end as BrutoBalanca,                                                             " & vbCrLf &
                 "    case                                                                             " & vbCrLf &
                 "	when isnull(R.Processo,'') = ''                                                    " & vbCrLf &
                 "		then isnull(P.Liquido,0)                                                       " & vbCrLf &
                 "		else isnull(R.PesoLiquido,0)                                                   " & vbCrLf &
                 "    end as Liquido,                                                                  " & vbCrLf &
                 "    ISNULL(nfXi.QuantidadeFiscal, 0) as 'Qtde Fiscal',                               " & vbCrLf &
                 "    ISNULL(NF.Nota_Id, 0) AS 'N° Nota', isnull(NFXI.Valor,0) as 'Valor Nota', Transp.Cliente_Id as 'CNPJ Transportadora', UPPER(TRANSP.Inscricao) as	'IE Transportadora'," & vbCrLf &
                 "    TRANSP.Nome AS 'Nome Transportador', isnull(CRTC.Nota_Id, 0) as 'Num Crtc', ISNULL(CRTCXI.Valor,0) AS 'Valor Crtc', " & vbCrLf &
                 " isnull(convert(varchar,CRTC.Movimento,103),'') as 'Movimento Crtc'," & vbCrLf &
                 " case" & vbCrLf &
                 "     when Provisao = 1" & vbCrLf &
                 "         then isnull(convert(varchar,CP.Baixa,103),'')" & vbCrLf &
                 "         else isnull(convert(varchar,CP.Prorrogacao,103),'')" & vbCrLf &
                 "     end as 'Vencimento Crtc'," & vbCrLf &
                 " case" & vbCrLf &
                 "     when Provisao = 1" & vbCrLf &
                 "         then 'PAGO'" & vbCrLf &
                 "         else ''" & vbCrLf &
                 "     end as 'Financeiro'," & vbCrLf &
                 " Pd.FreteCIFFOB, " & vbCrLf &
                 "    ISNULL(PL.CpfMotorista,'') as 'CPF Motorista', " & vbCrLf &
                 "    ISNULL(CLP.Nome + ' - ' + CLP.Cidade + '/' + CLP.Estado,'') as 'Nome Motorista'" & vbCrLf &
                 "into #temp " & vbCrLf &
                  "	FROM PESAGEM P	" & vbCrLf &
                  "INNER JOIN PEDIDOS PD" & vbCrLf &
                  "	ON P.EMPRESA_ID = PD.EMPRESA_ID	" & vbCrLf &
                  "		AND P.ENDEMPRESA_ID = PD.ENDEMPRESA_ID" & vbCrLf &
                  "		AND P.PEDIDO = PD.PEDIDO_ID" & vbCrLf &
                  "LEFT JOIN ROMANEIOSXPESAGENS RXP" & vbCrLf &
                  "	ON		P.EMPRESA_ID = RXP.EMPRESA_ID " & vbCrLf &
                  "		AND P.ENDEMPRESA_ID = RXP.ENDEMPRESA_ID" & vbCrLf &
                  "		AND P.PESAGEM_ID = RXP.PESAGEM_ID" & vbCrLf &
                  "		AND P.SEQUENCIA_ID = RXP.SEQUENCIA_ID" & vbCrLf &
                  "LEFT JOIN ROMANEIOS R" & vbCrLf &
                  "	ON		R.EMPRESA_ID = RXP.EMPRESA_ID" & vbCrLf &
                  "		AND R.ENDEMPRESA_ID = RXP.ENDEMPRESA_ID" & vbCrLf &
                  "		AND R.ROMANEIO_ID = RXP.ROMANEIO_ID" & vbCrLf &
                  "LEFT JOIN NOTASFISCAISXROMANEIOS NFXR" & vbCrLf &
                  "	ON		NFXR.EMPRESA_ID = R.EMPRESA_ID" & vbCrLf &
                  "		AND NFXR.ENDEMPRESA_ID = R.ENDEMPRESA_ID" & vbCrLf &
                  "		AND NFXR.ROMANEIO_ID = R.ROMANEIO_ID" & vbCrLf &
                  "LEFT JOIN NOTASFISCAISXITENS NFXI" & vbCrLf &
                  "	ON		NFXI.EMPRESA_ID = NFXR.EMPRESA_ID" & vbCrLf &
                  "		AND NFXI.ENDEMPRESA_ID = NFXR.ENDEMPRESA_ID" & vbCrLf &
                  "		AND	NFXI.CLIENTE_ID = NFXR.CLIENTE_ID" & vbCrLf &
                  "		AND	NFXI.ENDCLIENTE_ID = NFXR.ENDCLIENTE_ID" & vbCrLf &
                  "		AND NFXI.ENTRADASAIDA_ID = NFXR.ENTRADASAIDA_ID" & vbCrLf &
                  "		AND NFXI.SERIE_ID = NFXR.SERIE_ID" & vbCrLf &
                  "		AND NFXI.NOTA_ID = NFXR.NOTA_ID" & vbCrLf &
                  "LEFT JOIN NOTASFISCAIS NF" & vbCrLf &
                  "	ON		NF.EMPRESA_ID = NFXI.EMPRESA_ID" & vbCrLf &
                  "		AND NF.ENDEMPRESA_ID = NFXI.ENDEMPRESA_ID" & vbCrLf &
                  "		AND	NF.CLIENTE_ID = NFXI.CLIENTE_ID" & vbCrLf &
                  "		AND	NF.ENDCLIENTE_ID = NFXI.ENDCLIENTE_ID" & vbCrLf &
                  "		AND NF.ENTRADASAIDA_ID = NFXI.ENTRADASAIDA_ID" & vbCrLf &
                  "		AND NF.SERIE_ID = NFXI.SERIE_ID" & vbCrLf &
                  "		AND NF.NOTA_ID = NFXI.NOTA_ID" & vbCrLf &
                  "LEFT JOIN NOTASXNOTAS NXN" & vbCrLf &
                  "	ON		NXN.OrigemEmpresa_Id = NFXR.EMPRESA_ID" & vbCrLf &
                  "		AND NXN.OrigemEndEmpresa_Id = NFXR.ENDEMPRESA_ID" & vbCrLf &
                  "		AND NXN.OrigemCliente_Id  = NFXR.CLIENTE_ID" & vbCrLf &
                  "		AND NXN.OrigemEndCliente_Id = NFXR.ENDCLIENTE_ID" & vbCrLf &
                  "		AND NXN.OrigemEntradaSaida_Id = NFXR.ENTRADASAIDA_ID" & vbCrLf &
                  "		AND NXN.OrigemSerie_Id = NFXR.SERIE_ID" & vbCrLf &
                  "		AND NXN.OrigemNota_Id = NFXR.NOTA_ID" & vbCrLf &
                  "LEFT JOIN NOTASFISCAIS CRTC" & vbCrLf &
                  "	ON		CRTC.Empresa_Id = NXN.Empresa_Id" & vbCrLf &
                  "		AND CRTC.EndEmpresa_Id = NXN.EndEmpresa_Id " & vbCrLf &
                  "		AND CRTC.CLIENTE_ID = NXN.CLIENTE_ID " & vbCrLf &
                  "		AND CRTC.ENDCLIENTE_ID   = NXN.ENDCLIENTE_ID" & vbCrLf &
                  "		AND CRTC.ENTRADASAIDA_ID = NXN.ENTRADASAIDA_ID" & vbCrLf &
                  "		AND CRTC.Serie_Id = NXN.Serie_Id" & vbCrLf &
                  "		AND CRTC.Nota_Id = NXN.Nota_Id" & vbCrLf &
                  "		AND CRTC.TipoDeDocumento = 2" & vbCrLf &
                  "LEFT JOIN NOTASFISCAISXITENS CRTCXI" & vbCrLf &
                  "	ON			CRTCXI.EMPRESA_ID = CRTC.EMPRESA_ID" & vbCrLf &
                  "			AND CRTCXI.ENDEMPRESA_ID = CRTC.ENDEMPRESA_ID" & vbCrLf &
                  "			AND	CRTCXI.CLIENTE_ID = CRTC.CLIENTE_ID" & vbCrLf &
                  "			AND	CRTCXI.ENDCLIENTE_ID = CRTC.ENDCLIENTE_ID" & vbCrLf &
                  "			AND CRTCXI.ENTRADASAIDA_ID = CRTC.ENTRADASAIDA_ID" & vbCrLf &
                  "			AND CRTCXI.SERIE_ID = CRTC.SERIE_ID" & vbCrLf &
                  "			AND CRTCXI.NOTA_ID = CRTC.NOTA_ID" & vbCrLf &
                  " LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf &
                  "	ON			NFXT.Empresa_Id = CRTC.Empresa_Id            " & vbCrLf &
                  "			AND NFXT.EndEmpresa_Id = CRTC.EndEmpresa_Id      " & vbCrLf &
                  "			AND NFXT.Cliente_Id = CRTC.Cliente_Id            " & vbCrLf &
                  "			AND NFXT.EndCliente_Id = CRTC.EndCliente_Id      " & vbCrLf &
                  "			AND NFXT.EntradaSaida_Id = CRTC.EntradaSaida_Id  " & vbCrLf &
                  "			AND NFXT.Serie_Id = CRTC.Serie_Id                " & vbCrLf &
                  "			AND NFXT.Nota_Id = CRTC.Nota_Id                  " & vbCrLf

        If FinanceiroNovo Then
            sql &= "LEFT JOIN Titulos CP" & vbCrLf &
                 "   ON CP.Titulo_Id = NFXT.Titulo_Id" & vbCrLf &
                 "   AND CP.RecPag = 'P'" & vbCrLf
        Else
            sql &= "LEFT JOIN ContasAPagar CP                                   " & vbCrLf &
                  "	ON			NFXT.Titulo_Id = CP.Registro_Id              " & vbCrLf
        End If

        sql &= "INNER JOIN CLIENTES CL" & vbCrLf &
                  "	ON P.CLIENTE = CL.CLIENTE_ID" & vbCrLf &
                  "		AND P.ENDCLIENTE = CL.ENDERECO_ID" & vbCrLf &
                  "LEFT JOIN CLIENTES TRANSP" & vbCrLf &
                  "	ON TRANSP.CLIENTE_ID = P.TRANSPORTADOR" & vbCrLf &
                  "		AND TRANSP.Endereco_Id = P.ENDTRANSPORTADOR" & vbCrLf &
                  "INNER JOIN Clientes EMP" & vbCrLf &
                  "	ON EMP.Cliente_Id = P.Empresa_Id" & vbCrLf &
                  "	AND EMP.Endereco_Id = P.EndEmpresa_Id" & vbCrLf &
                  "INNER JOIN PRODUTOS PR" & vbCrLf &
                  "	ON P.PRODUTO = PR.PRODUTO_ID	" & vbCrLf &
                  "LEFT JOIN PLACAS PL	" & vbCrLf &
                  "	ON P.PLACA = PL.PLACA_ID" & vbCrLf &
                  "LEFT JOIN CLIENTES CLP " & vbCrLf &
                  "	ON PL.CPFMOTORISTA = CLP.CLIENTE_ID" & vbCrLf &
                  "	AND PL.ENDCPFMOTORISTA = CLP.ENDERECO_ID" & vbCrLf &
                  "INNER JOIN  Clientes AS ClientesXU" & vbCrLf &
                  "           ON PD.UnidadeDeNegocio = ClientesXU.Cliente_Id" & vbCrLf &
                                "	WHERE (P.Situacao = 1) AND (P.Sequencia_Id = 0)" & vbCrLf

        If ddlUnidadeDeNegocio.SelectedIndex > 0 Then
            sql &= " AND (PD.UnidadeDeNegocio = '" & ddlUnidadeDeNegocio.SelectedValue & "') " & vbCrLf
        End If

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            sql &= " AND (P.Empresa_Id = '" & Empresa(0) & "' AND P.EndEmpresa_Id = " & Empresa(1) & ") " & vbCrLf
        End If

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            sql &= " AND (P.Cliente = '" & Cliente(0) & "' AND P.EndCliente = " & Cliente(1) & ") " & vbCrLf
        End If

        If txtCodigoTransportador.Value.ToString.Length > 0 Then
            Dim Transportador() As String = txtCodigoTransportador.Value.ToString.Split("-")
            sql &= " AND (P.Transportador = '" & Transportador(0) & "' AND P.EndTransportador = " & Transportador(1) & ") " & vbCrLf
        End If

        If TxtPlaca.Text.Length > 0 Then
            If chkTodas.Checked Then
                sql &= " AND (P.Placa = '" & LTrim(RTrim(TxtPlaca.Text)) & "' or PL.Placa01 = '" & LTrim(RTrim(TxtPlaca.Text)) & "' or PL.Placa02 = '" & LTrim(RTrim(TxtPlaca.Text)) & "' or PL.Placa03 = '" & LTrim(RTrim(TxtPlaca.Text)) & "') " & vbCrLf
            Else
                sql &= " AND (P.Placa = '" & LTrim(RTrim(TxtPlaca.Text)) & "') " & vbCrLf
            End If
        End If

        If txtDataInicial.Text <> "" And txtDataFinal.Text <> "" Then
            sql &= " AND (P.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
        End If

        sql &= "update #temp set" & vbCrLf &
            "     PrimeiraPesagem  = 0," & vbCrLf &
            " segundapesagem = 0" & vbCrLf &
            " where(row > 1)" & vbCrLf &
            "select * from #temp" & vbCrLf &
            "drop table #temp" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "RelatorioDeLaudosPorPlacas")

    End Function

    Function getDataset() As DataSet
        Dim Sql As String

        Sql = "SELECT ClientesXU.Nome AS UnidadeDeNegocio, Pesagem.Empresa_Id, ClientesXE.Nome AS NomeEmpresa, " & vbCrLf &
              "		  Pesagem.Pesagem_Id, CONVERT(varchar, Pesagem.Movimento, 103) AS Movimento, Pesagem.Pedido, " & vbCrLf &
              "		  Produtos.Nome AS NomeProduto, Clientes.Nome AS NomeCliente, Pesagem.Placa, Pesagem.EntradaSaida, " & vbCrLf &
              "		  Pesagem.PrimeiraPesagem, Pesagem.SegundaPesagem,  Pesagem.BrutoBalanca, Pesagem.Liquido, " & vbCrLf &
              "       Mot.nome as NomeMotorista, mot.cidade as CidadeMotorista, mot.estado as EstadoMotorista, Placas.CpfMotorista," & vbCrLf &
              "       Pesagem.EntradaBalanca, Pesagem.SaidaBalanca, Placas.Placa01, Placas.Placa02, Placas.Placa03 " & vbCrLf &
              "  FROM Pesagem " & vbCrLf &
              "	INNER JOIN Pedidos " & vbCrLf &
              "	   ON Pesagem.Empresa_Id     = Pedidos.Empresa_Id " & vbCrLf &
              "	  AND Pesagem.EndEmpresa_Id = Pedidos.EndEmpresa_Id " & vbCrLf &
              "	  AND Pesagem.Pedido        = Pedidos.Pedido_Id " & vbCrLf &
              "	INNER JOIN Produtos " & vbCrLf &
              "	   ON Pesagem.Produto = Produtos.Produto_Id " & vbCrLf &
              "	INNER JOIN  Clientes " & vbCrLf &
              "	   ON Pesagem.Cliente     = Clientes.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndCliente = Clientes.Endereco_Id " & vbCrLf &
              "	INNER JOIN  Clientes AS ClientesXE " & vbCrLf &
              "	   ON Pesagem.Empresa_Id     = ClientesXE.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndEmpresa_Id = ClientesXE.Endereco_Id " & vbCrLf &
              "	INNER JOIN  Clientes AS ClientesXU " & vbCrLf &
              "	   ON Pedidos.UnidadeDeNegocio = ClientesXU.Cliente_Id " & vbCrLf &
              "	  AND 0                        = ClientesXU.Endereco_Id" & vbCrLf &
              "	INNER JOIN Placas " & vbCrLf &
              "	   ON Pesagem.Placa        = Placas.Placa_Id " & vbCrLf &
              " Inner Join Clientes Mot" & vbCrLf &
              "    on mot.Cliente_id  = placas.cpfMotorista" & vbCrLf &
              "   and mot.endereco_id = placas.endcpfmotorista" & vbCrLf &
              " WHERE Pesagem.Situacao     = 1 " & vbCrLf &
              "   AND Pesagem.Sequencia_Id = 0"

        If ddlUnidadeDeNegocio.SelectedIndex > 0 Then
            Sql &= " AND (Pedidos.UnidadeDeNegocio = '" & ddlUnidadeDeNegocio.SelectedValue & "') " & vbCrLf
        End If

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Sql &= " AND (Pesagem.Empresa_Id = '" & Empresa(0) & "' AND Pesagem.EndEmpresa_Id = " & Empresa(1) & ") " & vbCrLf
        End If

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= " AND (Pesagem.Cliente = '" & Cliente(0) & "' AND Pesagem.EndCliente = " & Cliente(1) & ") " & vbCrLf
        End If

        If txtCodigoTransportador.Value.ToString.Length > 0 Then
            Dim Transportador() As String = txtCodigoTransportador.Value.ToString.Split("-")
            Sql &= " AND (Pesagem.Transportador = '" & Transportador(0) & "' AND Pesagem.EndTransportador = " & Transportador(1) & ") " & vbCrLf
        End If

        If TxtPlaca.Text.Length > 0 Then
            If chkTodas.Checked Then
                Sql &= " AND (Pesagem.Placa = '" & LTrim(RTrim(TxtPlaca.Text)) & "' or Placas.Placa01 = '" & LTrim(RTrim(TxtPlaca.Text)) & "' or Placas.Placa02 = '" & LTrim(RTrim(TxtPlaca.Text)) & "' or Placas.Placa03 = '" & LTrim(RTrim(TxtPlaca.Text)) & "') " & vbCrLf
            Else
                Sql &= " AND (Pesagem.Placa = '" & LTrim(RTrim(TxtPlaca.Text)) & "') " & vbCrLf
            End If
        End If

        If txtDataInicial.Text <> "" And txtDataFinal.Text <> "" Then
            Sql &= " AND (Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
        End If

        Sql &= "Order by Pesagem.Movimento, Pesagem.Pesagem_Id"

        Return Banco.ConsultaDataSet(Sql, "RelatorioDeLaudosPorPlacas")
    End Function

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioDeLaudosPorPlacas", "RELATORIO") Then
                EmitirRelatorio(True)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioDeLaudosPorPlacas", "RELATORIO") Then
                EmitirRelatorio(False)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal pdf As Boolean)
        Dim ds As New DataSet
        Dim dsExcel As New DataSet
        Dim objEmpresa As [Lib].Negocio.Cliente

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            objEmpresa = New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
        Else
            objEmpresa = New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        End If

        If Not pdf Then
            dsExcel = GetDatasetExcel()

            If dsExcel Is Nothing OrElse dsExcel.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Placa sem movimentação!")
            Else
                Try
                    For Each row As DataRow In dsExcel.Tables(0).Rows
                        row("CNPJ Empresa") = Funcoes.FormatarCpfCnpj(row("CNPJ Empresa"))
                        row("CNPJ Transportadora") = Funcoes.FormatarCpfCnpj(row("CNPJ Transportadora"))
                        row("CPF Motorista") = Funcoes.FormatarCpfCnpj(row("CPF Motorista"))
                    Next
                    GerarExcel(dsExcel, objEmpresa)
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            End If
        Else
            ds = getDataset()

            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Placa sem Movimentação.")
            Else
                For Each row As DataRow In ds.Tables(0).Rows
                    row("Empresa_Id") = Funcoes.FormatarCpfCnpj(row("Empresa_Id"))
                    row("CpfMotorista") = Funcoes.FormatarCpfCnpj(row("CpfMotorista"))
                Next

                Dim param As New Dictionary(Of String, Object)
                param.Add("Titulo", "Relatorio de Laudos por Placa")
                param.Add("Nome", objEmpresa.Nome)
                param.Add("Cidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)

                Funcoes.BindReport(Me.Page, ds, IIf(ddlEmpresa.SelectedIndex > 0, "Cr_RelatorioDeLaudosPorPlacas", "Cr_RelatorioDeLaudosPorPlacasEmpresa"), eExportType.PDF, param)
            End If
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            ddlUnidadeDeNegocio.SelectedIndex = 0
            ddlEmpresa.Items.Clear()
            TxtPlaca.Text = ""
            txtCliente.Text = ""
            txtCodigoCliente.Value = ""
            txtTransportador.Text = ""
            txtCodigoTransportador.Value = ""
            chkTodas.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeLaudosPorPlacas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class