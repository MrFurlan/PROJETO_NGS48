Imports System.IO
Imports System.Data
Imports System.Drawing
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class RelatorioDeTitulos
    Inherits BasePage
    Dim EmpresasConsolidas As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeTitulos", "ACESSAR") Then
                    txtPeriodoInicialConsultaTitulos.Text = "01/" & Format(DateTime.Now.Month, "00") & "/" & DateTime.Now.Year
                    txtPeriodoFinalConsultaTitulos.Text = Format(DateTime.Now, "dd/MM/yyyy")
                    ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "")
                    ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
                    CarteiraDoTitulo()
                    CargaCarteiras()
                    CarregarBancos()
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                    ddl.Carregar(ddlRepresentante, CarregarDDL.Tabela.ClientesXTipos, "6")
                    ddl.Carregar(lstFinalidade, CarregarDDL.Tabela.Finalidade, "", False)

                    Dim Parametros As New Hashtable
                    Parametros.Clear()
                    Parametros.Add("listarTudo", "N")

                    ddl.Carregar(lstTipoPagRec, CarregarDDL.Tabela.TipoDePagamento, "", False, Parametros)

                    RadAtivos.Checked = True
                    HID.Value = Guid.NewGuid().ToString()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue, True)
    End Sub

    Private Sub CarregarBancos()
        Dim sql = "SELECT DISTINCT Bancos.Banco_Id,  Bancos.Descricao" & vbCrLf &
                  "  FROM BancosXContas " & vbCrLf &
                  " INNER JOIN Bancos  " & vbCrLf &
                  "    ON BancosXContas.Banco_Id = Bancos.Banco_Id " & vbCrLf &
                  " ORDER BY Banco_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Bancos")

        ddlBanco.Items.Clear()

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                ddlBanco.Items.Add(New ListItem(Format(row("Banco_Id"), "0000") & " - " & row("Descricao"), row("Banco_Id")))
            Next
        End If

        ddlBanco.Items.Insert(0, "")
    End Sub

    Private Sub CargaCarteiras()
        ddl.Carregar(ddlFinalidadeFinanceira, CarregarDDL.Tabela.CarteiraFinanceira, " Classificacao =" & IIf(RdReceber.Checked, "'R'", "'P'"), True)
    End Sub

    Private Sub CarteiraDoTitulo()
        Dim objCarteiraDoTitulo As New [Lib].Negocio.ListCarteiraDoTitulo()

        ddlCarteiraDoTitulo.DataValueField = "Codigo"
        ddlCarteiraDoTitulo.DataTextField = "Descricao"
        ddlCarteiraDoTitulo.DataSource = objCarteiraDoTitulo.ToArray()
        ddlCarteiraDoTitulo.DataBind()

        Funcoes.InserirLinhaEmBranco(ddlCarteiraDoTitulo)
    End Sub

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
        chkConsolidarCliente.Checked = False
        chkConsolidarEmpresa.Checked = False
        txtClientes.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ddlRepresentante.SelectedIndex = 0
        ddlCarteiraDoTitulo.SelectedIndex = 0
        ddlFinalidadeFinanceira.SelectedIndex = 0
        ddlBanco.SelectedIndex = 0
        ddlSafra.SelectedIndex = 0
        chkPeriodo.Checked = True
        rbDocumento.Enabled = True
        rbLiquido.Enabled = True
        ddlMoeda.SelectedIndex = 0
        txtPedido.Text = String.Empty
        txtPeriodoInicialConsultaTitulos.Text = "01/" & Format(DateTime.Now.Month, "00") & "/" & DateTime.Now.Year
        txtPeriodoFinalConsultaTitulos.Text = Format(DateTime.Now, "dd/MM/yyyy")
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objRelatorioTitulo" & HID.Value) IsNot Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objRelatorioTitulo" & HID.Value)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objRelatorioTitulo" & HID.Value)
        ElseIf Session("ObjPedidos" & HID.Value) IsNot Nothing Then
            txtPedido.Text = CType(Session("ObjPedidos" & HID.Value), [Lib].Negocio.Pedido).Codigo
            Session.Remove("ObjPedidos" & HID.Value)
        End If
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidadeConsultaTitulos.SelectedIndexChanged
        Try
            CarregarEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdLimparConsultaTitulos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RadTodos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkPrevisao.Checked = False
            chkProvisao.Checked = False
            rbLiquido.Enabled = False
            rbDocumento.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RadBaixados_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkPrevisao.Checked = False
            chkProvisao.Checked = False
            rbLiquido.Enabled = True
            rbDocumento.Enabled = True
            rbLiquido.Checked = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RadAtivos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkPrevisao.Visible Then
                chkPrevisao.Checked = True
                chkProvisao.Checked = False
            End If
            rbLiquido.Enabled = False
            rbDocumento.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdReceber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaCarteiras()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdPagar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaCarteiras()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub radDataBase_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If radDataBase.Checked = True Then
                chkProvisao.Checked = False
                rbDocumento.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedido.Click
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "Cliente não foi selecionado.")
            Else
                HttpContext.Current.Session("ssCampo" & HID.Value) = ""
                HttpContext.Current.Session("ssTipoRetorno") = "ObjPedidos" & HID.Value
                Dim parameters As New Dictionary(Of String, Object)
                parameters("unidade") = DdlUnidadeConsultaTitulos.SelectedValue
                parameters("empresa") = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0)
                parameters("enderecoEmpresa") = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1)
                parameters("cliente") = txtCodigoCliente.Value.Split("-")(0)
                parameters("enderecoCliente") = txtCodigoCliente.Value.Split("-")(1)
                ucConsultaPedidos.Limpar()
                ucConsultaPedidos.BindGridView(parameters)
                Popup.ConsultaDePedidos(Me.Page, "ObjPedidos" & HID.Value)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objRelatorioTitulo" & HID.Value, "txtNome")
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkPeriodo.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub radData_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkListarProdutos.Visible = False
            chkListarProdutos.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub radClienteData_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkListarProdutos.Visible = True
            chkListarProdutos.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdClientePedido_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkListarProdutos.Visible = True
            chkListarProdutos.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub listarEmpresas()
        Dim sqlConsulta As String = "SELECT Empresa_Id FROM ClientesXEmpresas where left(Empresa_Id,8) = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0).Substring(0, 8) & "'"
        Dim dsEmpesa As New DataSet
        dsEmpesa = Banco.ConsultaDataSet(sqlConsulta, "Empresas")

        Dim primeiro As Boolean = True

        For Each Dr As DataRow In dsEmpesa.Tables(0).Rows
            If primeiro Then
                EmpresasConsolidas += "'" & CStr(Dr("Empresa_Id"))
                primeiro = False
            Else
                EmpresasConsolidas += "','" & CStr(Dr("Empresa_Id"))
            End If
        Next

        EmpresasConsolidas += "'"
    End Sub

    Private Function getDataSetMovtoBancario() As DataSet
        Dim sql As String
        sql = "SELECT Empresa, Conta," & vbCrLf &
              "       CASE WHEN ContaPartida IS NULL THEN ContaPartidaB ELSE ContaPartida END AS Partida," & vbCrLf &
              "       CASE WHEN ContaPartida IS NULL THEN ClientePartidaB ELSE ClientePartida END AS CNPJ," & vbCrLf &
              "       CASE WHEN ContaPartida IS NULL THEN NomePartidaB ELSE NomePartida END AS Nome," & vbCrLf &
              "       DebitoOficial AS Debito, CreditoOficial AS Credito, Historico, Lote, Titulo, Movimento, Vencimento, Prorrogacao, Baixa, Observacoes, Descricao AS Carteira," & vbCrLf &
              "       NomeEncargo AS Encargo, Pedido," & vbCrLf &
              "       (SELECT TOP (1) PxI.Produto_Id + ' - ' + Produtos.Nome AS Produto" & vbCrLf &
              "          FROM PedidoXItem AS PxI " & vbCrLf &
              "         INNER JOIN Produtos" & vbCrLf &
              "            ON PxI.Produto_Id = Produtos.Produto_Id" & vbCrLf &
              "         WHERE PxI.Empresa_Id    = Consulta.EmpresaPedido" & vbCrLf &
              "           AND PxI.EndEmpresa_Id = Consulta.EndEmpresaPedido" & vbCrLf &
              "           AND PxI.Pedido_Id     = Consulta.Pedido" & vbCrLf &
              "       ) AS Produto," & vbCrLf &
              "       isnull(Mestre, 'N') as Mestre, EmmpOri.Reduzido AS FilialOrigem, EmpresaOrigem" & vbCrLf &
              "  FROM (SELECT Empresa.Cliente_Id + '-' + Empresa.Reduzido + '-' + Empresa.Nome + '-' + Empresa.Cidade AS Empresa," & vbCrLf &
              "               Razao.Conta_Id + '-' + PlanoDeContas.Titulo AS Conta," & vbCrLf &
              "               CASE" & vbCrLf &
              "                 WHEN len(EncargosAPagar.ContaCredito) > 7" & vbCrLf &
              "                   THEN CASE" & vbCrLf &
              "                          WHEN Razao.creditoOficial > 0" & vbCrLf &
              "                            THEN EncargosAPagar.ContaDebito" & vbCrLf &
              "                            ELSE ContasAPagar.ContaContabilPagadora" & vbCrLf &
              "                        END" & vbCrLf &
              "                   ELSE ContasAPagar.ContaContabilCliente" & vbCrLf &
              "               END AS ContaPartida," & vbCrLf &
              "               CASE" & vbCrLf &
              "                 WHEN len(EncargosAPagar.ContaCredito) > 7" & vbCrLf &
              "                   THEN CASE" & vbCrLf &
              "                          WHEN Razao.creditoOficial > 0" & vbCrLf &
              "                            THEN ''" & vbCrLf &
              "                            ELSE ''" & vbCrLf &
              "                        END" & vbCrLf &
              "                   ELSE ContasAPagar.Cliente" & vbCrLf &
              "               END AS ClientePartida," & vbCrLf &
              "               CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                                      > 7 THEN CASE WHEN Razao.creditoOficial > 0 THEN ContasEncargosAPagar.Titulo ELSE PlanoDeContasPagadora.Titulo END ELSE Fornecedores.Nome END" & vbCrLf &
              "                                       AS NomePartida, CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                                      > 7 THEN CASE WHEN Razao.DebitoOficial > 0 THEN EncargosAReceber.ContaCredito ELSE ContasAReceber.ContaContabilPagadora END ELSE ContasAReceber.ContaContabilCliente" & vbCrLf &
              "                                       END AS ContaPartidaB, CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                                      > 7 THEN CASE WHEN Razao.DebitoOficial > 0 THEN '' ELSE '' END ELSE ContasAReceber.Cliente END AS ClientePartidaB," & vbCrLf &
              "                                      CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                                      > 7 THEN CASE WHEN Razao.DebitoOficial > 0 THEN ContasEncargosAReceber.Titulo ELSE PlanoDeContasRecebedora.Titulo END ELSE Clientes.Nome END" & vbCrLf &
              "                                       AS NomePartidaB, Razao.DebitoOficial, Razao.CreditoOficial, Razao.Lote_Id AS Lote, Razao.Titulo, ContasAPagar.Carteira," & vbCrLf &
              "                                      ComprasXProdutosAPagar.Descricao, ISNULL(ContasAPagar.Tributo, N'') AS Encargo, ISNULL(EncargosAPagar.Descricao, N'') AS NomeEncargo," & vbCrLf &
              "                                      ContasAPagar.Moeda, CASE WHEN ContasAPagar.UsuarioInclusaoData IS NULL" & vbCrLf &
              "                                      THEN ContasAReceber.UsuarioInclusaoData ELSE ContasAPagar.UsuarioInclusaoData END AS Movimento, CASE WHEN ContasAPagar.Vencimento IS NULL" & vbCrLf &
              "                                       THEN ContasAReceber.Vencimento ELSE ContasAPagar.Vencimento END AS Vencimento, CASE WHEN ContasAPagar.Prorrogacao IS NULL" & vbCrLf &
              "                                      THEN ContasAReceber.Prorrogacao ELSE ContasAPagar.Prorrogacao END AS Prorrogacao, Razao.Movimento_Id AS Baixa," & vbCrLf &
              "                                      CASE WHEN ContasAPagar.Pedido IS NULL THEN ContasAReceber.EmpresaPedido ELSE ContasAPagar.EmpresaPedido END AS EmpresaPedido," & vbCrLf &
              "                                      CASE WHEN ContasAPagar.Pedido IS NULL" & vbCrLf &
              "                                      THEN ContasAReceber.EndEmpresaPedido ELSE ContasAPagar.EndEmpresaPedido END AS EndEmpresaPedido," & vbCrLf &
              "                                      CASE WHEN ContasAPagar.Pedido IS NULL THEN ContasAReceber.Pedido ELSE ContasAPagar.Pedido END AS Pedido, Razao.Historico," & vbCrLf &
              "                                      CONVERT(varchar, ContasAPagar.Observacoes) As Observacoes, " & vbCrLf &
              "                                      CASE WHEN Razao.creditoOficial > 0 THEN ContasAPagar.Grupado  ELSE ContasAPagar.Grupado  End as  Mestre, " & vbCrLf &
              "                                      CASE WHEN ContasAPagar.Empresa IS NULL THEN ContasAReceber.Empresa ELSE ContasAPagar.Empresa END AS EmpresaOrigem," & vbCrLf &
              "                                      CASE WHEN ContasAPagar.EndEmpresa IS NULL THEN ContasAReceber.EndEmpresa ELSE ContasAPagar.EndEmpresa END AS EndEmpresaOrigem" & vbCrLf &
              "               FROM          Razao INNER JOIN" & vbCrLf &
              "                                      PlanoDeContas AS PlanoDeContas ON Razao.Conta_Id = PlanoDeContas.Conta_Id INNER JOIN" & vbCrLf &
              "                                      Clientes AS Empresa ON Razao.Empresa_Id = Empresa.Cliente_Id AND Razao.EndEmpresa_Id = Empresa.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
              "                                      Clientes AS Clientes RIGHT OUTER JOIN" & vbCrLf &
              "                                      ComprasXProdutos RIGHT OUTER JOIN" & vbCrLf &
              "                                      PlanoDeContas AS PlanoDeContasRecebedora RIGHT OUTER JOIN" & vbCrLf &
              "                                      ContasAReceber ON PlanoDeContasRecebedora.Conta_Id = ContasAReceber.ContaContabilPagadora ON" & vbCrLf &
              "                                      ComprasXProdutos.Produto_Id = ContasAReceber.Carteira ON Clientes.Cliente_Id = ContasAReceber.Cliente AND" & vbCrLf &
              "                                      Clientes.Endereco_Id = ContasAReceber.EndCliente LEFT OUTER JOIN" & vbCrLf &
              "                                      PlanoDeContas AS ContasEncargosAReceber RIGHT OUTER JOIN" & vbCrLf &
              "                                      Encargos AS EncargosAReceber ON ContasEncargosAReceber.Conta_Id = EncargosAReceber.ContaCredito ON" & vbCrLf &
              "                                      ContasAReceber.Tributo = EncargosAReceber.Encargo_id ON Razao.Titulo = ContasAReceber.Registro_Id LEFT OUTER JOIN" & vbCrLf &
              "                                      PlanoDeContas AS PlanoDeContasPagadora RIGHT OUTER JOIN" & vbCrLf &
              "                                      ContasAPagar ON PlanoDeContasPagadora.Conta_Id = ContasAPagar.ContaContabilPagadora ON" & vbCrLf &
              "                                      Razao.Titulo = ContasAPagar.Registro_Id LEFT OUTER JOIN" & vbCrLf &
              "                                      Encargos AS EncargosAPagar LEFT OUTER JOIN" & vbCrLf &
              "                                      PlanoDeContas AS ContasEncargosAPagar ON EncargosAPagar.ContaCredito = ContasEncargosAPagar.Conta_Id ON" & vbCrLf &
              "                                      ContasAPagar.Tributo = EncargosAPagar.Encargo_id LEFT OUTER JOIN" & vbCrLf &
              "                                      ComprasXProdutos AS ComprasXProdutosAPagar ON ContasAPagar.Carteira = ComprasXProdutosAPagar.Produto_Id LEFT OUTER JOIN" & vbCrLf &
              "                                      Clientes AS Fornecedores ON ContasAPagar.Cliente = Fornecedores.Cliente_Id AND ContasAPagar.EndCliente = Fornecedores.Endereco_Id" & vbCrLf &
              " WHERE (Razao.Movimento_Id BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' AND '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "')" & vbCrLf &
              "   AND (Razao.Conta_Id LIKE '1010102%')) AS Consulta" & vbCrLf &
              "         INNER JOIN Clientes as EmmpOri " & vbCrLf &
              "                 on EmmpOri.Cliente_Id = EmpresaOrigem" & vbCrLf &
              "                 and EmmpOri.Endereco_Id = EndEmpresaOrigem" & vbCrLf &
              "   Order by Consulta.Empresa, Consulta.Conta, Consulta.Baixa, Consulta.Lote" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "MovtoBancario")
    End Function

    Private Sub GerarExcel()
        Dim rowIndex As Integer = 0
        Dim columnIndex As Integer = 0
        Try

            'ACRESCENTADO PARA CONSOLIDAR EMPRESA ONDE MELHORA A PERFORMANCE DO SQL - FURLAN - 11/10/2024
            If chkConsolidarEmpresa.Checked Then listarEmpresas()

            'Dim ds As DataSet = getDataSetExcelNew() 'getDataSetExcel()
            Dim ds As DataSet = getDataSetExcel()
            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("TÍTULOS")

                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        columnIndex += 1
                        'If col.ColumnName <> "Nota" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        '    'Else
                        '    '    'criando cabeçalho de coluna calculada
                        '    '    worksheet.Cells(rowIndex, columnIndex).Value = "Qtde Paga"
                        '    '    columnIndex += 1
                        '    '    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        'End If
                    Next

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count + 1)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                        range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                    End Using

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        rowIndex += 1
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            If col.ColumnName <> "Nota" Then
                                If col.ColumnName = "Empresa" OrElse col.ColumnName = "GrupodeContas" OrElse col.ColumnName = "Fornecedor" OrElse col.ColumnName = "Produto" _
                                    OrElse col.ColumnName = "ContaContabilEncargo" OrElse col.ColumnName = "Nota" OrElse col.ColumnName = "ContratoBanco" Then
                                    If Not IsDBNull(row(col.ColumnName)) AndAlso IsNumeric(row(col.ColumnName)) Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    End If
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                End If
                            Else

                                'If (Not String.IsNullOrWhiteSpace(row("Total")) AndAlso Not String.IsNullOrWhiteSpace(row("Quantidade")) AndAlso Not String.IsNullOrWhiteSpace(row("ValorLiquido"))) _
                                '    AndAlso (row("Total") <> 0 AndAlso row("Quantidade") <> 0 AndAlso row("ValorLiquido") <> 0) Then
                                '    'adicionando coluna calculada
                                '    worksheet.Cells(rowIndex, columnIndex).Formula = String.Format("=(AL{0} / AJ{0}) * AC{0}", rowIndex)
                                'Else
                                '    worksheet.Cells(rowIndex, columnIndex).Value = 0
                                'End If

                                'worksheet.Cells(rowIndex, columnIndex).Value = 0

                                'adiciona o numero da nota
                                'columnIndex += 1
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If
                            columnIndex += 1
                        Next

                        'aplicando formatação decimal nos campos de valores
                        worksheet.Cells(String.Format("AL{0}:AM{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        worksheet.Cells(String.Format("Z{0}:AE{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                        worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("AN{0}", rowIndex)).Style.Numberformat.Format = "0"
                        worksheet.Cells(String.Format("AO{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.Numberformat.Format = "0"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count + 1)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A1:AR" & rowIndex).AutoFilter = True

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
        Catch ex As Exception
            Throw New Exception(ex.Message & "  linha: " & rowIndex & " Coluna: " & columnIndex)
            MsgBox(Me.Page, rowIndex)

        End Try
    End Sub

    Public Sub GerarExcelMovtoBancario(ByVal ds As DataSet, ByVal titulo As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing)
        If ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado.")
        Else
            Dim rowIndex As Integer = 0
            Dim columnIndex As Integer = 0
            Try
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)

                        'criando aba da planilha.
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("TÍTULOS")

                        rowIndex += 1

                        'Inserindo o cabeçalho.
                        For Each col As DataColumn In ds.Tables(0).Columns
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        Next

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count + 1)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
                        End Using

                        ' Exportando conteúdo da planilha com os dados da Tabela.
                        For Each row As DataRow In ds.Tables(0).Rows
                            rowIndex += 1
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If col.ColumnName = "FilialOrigem" OrElse col.ColumnName = "EmpresaOrigem" OrElse col.ColumnName = "Partida" OrElse col.ColumnName = "CNPJ" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                ElseIf IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                End If

                                'Formatacoes de valores
                                If colunas IsNot Nothing Then
                                    For Each coluna In colunas
                                        If coluna.Key = col.ColumnName Then
                                            If coluna.Value = eTipoCampo.Numerico Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                                Exit For
                                            ElseIf coluna.Value = eTipoCampo.ValorComTotalizador OrElse coluna.Value = eTipoCampo.ValorSemTotalizador Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                Exit For
                                            ElseIf coluna.Value = eTipoCampo.Data Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                                columnIndex += 1
                            Next

                            'Formatacoes de celulas
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count + 1)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                        Next

                        'Soma dos campos de valores
                        If colunas IsNot Nothing Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count + 1)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            columnIndex = 1

                            For Each col In colunas
                                If col.Value = eTipoCampo.ValorComTotalizador Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                                    Exit For
                                End If
                            Next

                            For Each col As DataColumn In ds.Tables(0).Columns
                                For Each coluna In colunas
                                    If coluna.Key = col.ColumnName Then
                                        If coluna.Value = eTipoCampo.ValorComTotalizador Then
                                            worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(1, columnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, columnIndex).Address & ")"
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                            Exit For
                                        End If
                                    End If
                                Next
                                columnIndex += 1
                            Next
                        End If

                        'criando auto filtro na planilha
                        worksheet.Cells(1, 1, 1, columnIndex - 1).AutoFilter = True

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linham
                        worksheet.View.FreezePanes(2, 1)

                        CriaSegundaAba(worksheet, package, ds, colunas)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message & " - linha: " & rowIndex & " - coluna: " & columnIndex)
            End Try
        End If
    End Sub

    Private Sub CriaSegundaAba(ByRef worksheet As ExcelWorksheet, ByRef package As ExcelPackage, ByVal ds As DataSet, ByVal colunas As Dictionary(Of String, eTipoCampo))
        Dim rowIndex As Integer = 0
        Dim columnIndex As Integer = 0

        Dim Empresa As String = ""
        Dim Conta As String = ""
        Dim Saldo As Double

        worksheet = package.Workbook.Worksheets.Add("Razão")

        rowIndex += 1

        'Inserindo o cabeçalho.
        For Each col As DataColumn In ds.Tables(0).Columns
            columnIndex += 1
            If col.ColumnName = "Credito" Then
                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                columnIndex += 1
                worksheet.Cells(rowIndex, columnIndex).Value = "Saldo"
            Else
                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
            End If
        Next

        'aplicando formatação nas células do cabeçalho
        Using range = worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count + 1)
            range.Style.Font.Bold = True
            range.Style.Fill.PatternType = ExcelFillStyle.Solid
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
            range.Style.Font.Color.SetColor(Color.White)
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
            range.Style.Border.Bottom.Color.SetColor(Color.FromArgb(0, 0, 0))
        End Using

        ' Exportando conteúdo da planilha com os dados da Tabela.
        For Each row As DataRow In ds.Tables(0).Rows
            columnIndex = 1
            For Each col As DataColumn In ds.Tables(0).Columns
                If col.ColumnName = "Empresa" AndAlso row(col.ColumnName) <> Empresa Then
                    rowIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                    Empresa = row(col.ColumnName)

                    'worksheet.Cells(rowIndex, 1, rowIndex, 2).Merge = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True

                ElseIf col.ColumnName = "Empresa" AndAlso row(col.ColumnName) = Empresa Then
                    worksheet.Cells(rowIndex, columnIndex).Value = ""

                ElseIf col.ColumnName = "Conta" AndAlso row(col.ColumnName) <> Conta Then
                    rowIndex += 1

                    'worksheet.Cells(rowIndex, 2, rowIndex, 4).Merge = True
                    Using range = worksheet.Cells("B" & rowIndex & ":H" & rowIndex)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.Black)
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using

                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)

                    Conta = row(col.ColumnName)

                    Dim sql As String = String.Empty
                    sql = " select isnull(Sum(DebitoOficial - CreditoOficial), 0) as Saldo" & vbCrLf &
                        " from razao" & vbCrLf &
                        " where Empresa_Id = '" & Left(Empresa, 14) & "' And Conta_ID = '" & Left(Conta, 9) & "'" & vbCrLf &
                        " and Movimento_ID < '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf

                    Dim dsSaldo As DataSet = Banco.ConsultaDataSet(sql, "Saldo")

                    If dsSaldo IsNot Nothing AndAlso dsSaldo.Tables IsNot Nothing AndAlso dsSaldo.Tables(0).Rows.Count > 0 Then
                        worksheet.Cells(rowIndex, columnIndex + 3).Value = "SALDO INICIAL......................"
                        worksheet.Cells(rowIndex, columnIndex + 6).Value = dsSaldo.Tables(0).Rows(0)(0)
                        Saldo = dsSaldo.Tables(0).Rows(0)(0)
                        worksheet.Cells(rowIndex, columnIndex + 6).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        Using range = worksheet.Cells(rowIndex, columnIndex + 6)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Font.Color.SetColor(Color.Red)
                        End Using
                    End If
                    rowIndex += 1
                ElseIf col.ColumnName = "Conta" AndAlso row(col.ColumnName) = Conta Then
                    worksheet.Cells(rowIndex, columnIndex).Value = ""

                ElseIf col.ColumnName = "Credito" Then
                    worksheet.Cells(rowIndex, columnIndex - 1).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    columnIndex += 1

                    Saldo = Saldo + CDec(row("Debito")) - CDec(row("Credito"))
                    worksheet.Cells(rowIndex, columnIndex).Value = Saldo
                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    Using range = worksheet.Cells(rowIndex, columnIndex)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.None
                        range.Style.Font.Color.SetColor(Color.Red)
                    End Using

                ElseIf col.ColumnName = "FilialOrigem" OrElse col.ColumnName = "EmpresaOrigem" Then
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                ElseIf IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"

                ElseIf col.ColumnName = "Mestre" Then
                    If row(col.ColumnName) = "M" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = "Mestre"

                        Using range = worksheet.Cells(rowIndex, 3, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(191, 191, 191))
                        End Using
                    Else
                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                    End If
                Else
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                End If

                'Formatacoes de valores
                If colunas IsNot Nothing Then
                    For Each coluna In colunas
                        If coluna.Key = col.ColumnName Then
                            If coluna.Value = eTipoCampo.Numerico Then
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                Exit For
                            ElseIf coluna.Value = eTipoCampo.ValorComTotalizador OrElse coluna.Value = eTipoCampo.ValorSemTotalizador Then
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                Exit For
                            ElseIf coluna.Value = eTipoCampo.Data Then
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                Exit For
                            End If
                        End If
                    Next
                End If
                columnIndex += 1
            Next

            If row("Mestre") = "M" Then
                Dim dsFilhos As DataSet = getDataSetFilhos(row("Titulo"))

                If dsFilhos IsNot Nothing AndAlso dsFilhos.Tables IsNot Nothing AndAlso dsFilhos.Tables(0).Rows.Count > 0 Then
                    rowIndex += 1
                    For Each rowFilho As DataRow In dsFilhos.Tables(0).Rows
                        columnIndex = 3
                        For Each colFilho As DataColumn In dsFilhos.Tables(0).Columns
                            If colFilho.ColumnName <> "Empresa" AndAlso colFilho.ColumnName <> "Conta" Then
                                If colFilho.ColumnName = "ContaPartida" OrElse colFilho.ColumnName = "ClientePartida" Then
                                    If IsNumeric(rowFilho(colFilho.ColumnName)) Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(rowFilho(colFilho.ColumnName))
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = rowFilho(colFilho.ColumnName)
                                    End If
                                ElseIf colFilho.ColumnName = "Credito" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = rowFilho(colFilho.ColumnName)
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    columnIndex += 1

                                    worksheet.Cells(rowIndex, columnIndex).Value = Saldo
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                    Using range = worksheet.Cells(rowIndex, columnIndex)
                                        range.Style.Font.Bold = True
                                        range.Style.Fill.PatternType = ExcelFillStyle.None
                                        range.Style.Font.Color.SetColor(Color.Red)
                                    End Using
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = rowFilho(colFilho.ColumnName)
                                End If
                                columnIndex += 1
                            End If
                        Next

                        Using range = worksheet.Cells(rowIndex, 3, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242))
                        End Using

                        rowIndex += 1
                    Next
                End If
            End If

            'Formatacoes de celulas
            If rowIndex Mod 2 = 0 Then
                Using range = worksheet.Cells(rowIndex, 3, rowIndex, ds.Tables(0).Columns.Count + 1)
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                End Using
            End If
            rowIndex += 1
        Next

        'criando auto filtro na planilha
        worksheet.Cells(1, 1, 1, columnIndex - 1).AutoFilter = True

        'setando autofit nas células da planilha
        worksheet.Cells.AutoFitColumns(0)

        worksheet.Column(1).Width = 19
        worksheet.Column(2).Width = 49

        worksheet.Column(3).Width = 10
        worksheet.Column(4).Width = 15
        worksheet.Column(5).Width = 40

        worksheet.Column(6).Width = 15
        worksheet.Column(7).Width = 15
        worksheet.Column(8).Width = 15

        worksheet.Column(9).Width = 63

        'congelando primeira linham
        worksheet.View.FreezePanes(2, 1)

        package.Workbook.Worksheets("Razão").View.TabSelected = True

    End Sub

    Private Function getDataSetFilhos(ByVal titulo As Integer) As DataSet
        Dim sql As String = String.Empty
        sql = " SELECT Empresa, Conta, ContaPartida, ClientePartida, NomePartida, DebitoOficial AS Debito, CreditoOficial AS Credito, Historico, Lote, Titulo, Movimento, Vencimento," & vbCrLf &
              "        Prorrogacao, Baixa, Observacoes, Descricao AS Carteura, NomeEncargo AS Encargo, Pedido," & vbCrLf &
              "        (SELECT TOP (1) PxI.Produto_Id + ' - ' + Produtos.Nome AS Produto" & vbCrLf &
              "           FROM PedidoXItem AS PxI " & vbCrLf &
              "          INNER JOIN Produtos " & vbCrLf &
              "             ON PxI.Produto_Id = Produtos.Produto_Id" & vbCrLf &
              "          WHERE (PxI.Empresa_Id = Consulta.EmpresaPedido) " & vbCrLf &
              "            AND (PxI.EndEmpresa_Id = Consulta.EndEmpresaPedido) " & vbCrLf &
              "            AND (PxI.Pedido_Id = Consulta.Pedido)" & vbCrLf &
              "         ) AS Produto" & vbCrLf &
              "   FROM (SELECT Empresa.Cliente_Id + '-' + Empresa.Reduzido + '-' + Empresa.Nome + '-' + Empresa.Cidade AS Empresa," & vbCrLf &
              "                  Razao.Conta_Id + '-' + PlanoDeContas.Titulo AS Conta, CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                  > 7 THEN CASE WHEN Razao.creditoOficial > 0 THEN EncargosAPagar.ContaDebito ELSE ContasAPagar.ContaContabilPagadora END ELSE ContasAPagar.ContaContabilCliente" & vbCrLf &
              "                   END AS ContaPartida, CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                  > 7 THEN CASE WHEN Razao.creditoOficial > 0 THEN '' ELSE '' END ELSE ContasAPagar.Cliente END AS ClientePartida," & vbCrLf &
              "                  CASE WHEN len(EncargosAPagar.ContaCredito)" & vbCrLf &
              "                  > 7 THEN CASE WHEN Razao.creditoOficial > 0 THEN ContasEncargosAPagar.Titulo ELSE PlanoDeContasPagadora.Titulo END ELSE Fornecedores.Nome END" & vbCrLf &
              "                   AS NomePartida, Razao.DebitoOficial, Razao.CreditoOficial, Razao.Lote_Id AS Lote, Razao.Titulo, ContasAPagar.Carteira," & vbCrLf &
              "                  ComprasXProdutosAPagar.Descricao, ISNULL(ContasAPagar.Tributo, N'') AS Encargo, ISNULL(EncargosAPagar.Descricao, N'') AS NomeEncargo," & vbCrLf &
              "                  ContasAPagar.Moeda, Razao.Movimento_Id AS Baixa, Razao.Historico, CONVERT(varchar, ContasAPagar.Observacoes) AS Observacoes," & vbCrLf &
              "                  ContasAPagar.UsuarioInclusaoData AS Movimento, ContasAPagar.Vencimento, ContasAPagar.Prorrogacao, ContasAPagar.Pedido," & vbCrLf &
              "                  ContasAPagar.EmpresaPedido , ContasAPagar.EndEmpresaPedido" & vbCrLf &
              "             FROM Razao " & vbCrLf &
              "            INNER JOIN PlanoDeContas AS PlanoDeContas " & vbCrLf &
              "               ON Razao.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf &
              "            INNER JOIN Clientes AS Empresa " & vbCrLf &
              "               ON Razao.Empresa_Id = Empresa.Cliente_Id " & vbCrLf &
              "              AND Razao.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf &
              "             LEFT OUTER JOIN PlanoDeContas AS PlanoDeContasPagadora " & vbCrLf &
              "            RIGHT OUTER JOIN ContasAPagar ON PlanoDeContasPagadora.Conta_Id = ContasAPagar.ContaContabilPagadora " & vbCrLf &
              "               ON Razao.Titulo = ContasAPagar.Registro_Id " & vbCrLf &
              "             LEFT OUTER JOIN Encargos AS EncargosAPagar " & vbCrLf &
              "             LEFT OUTER JOIN PlanoDeContas AS ContasEncargosAPagar " & vbCrLf &
              "               ON EncargosAPagar.ContaCredito = ContasEncargosAPagar.Conta_Id " & vbCrLf &
              "               ON ContasAPagar.Tributo = EncargosAPagar.Encargo_id " & vbCrLf &
              "             LEFT OUTER JOIN ComprasXProdutos AS ComprasXProdutosAPagar " & vbCrLf &
              "               ON ContasAPagar.Carteira = ComprasXProdutosAPagar.Produto_Id " & vbCrLf &
              "             LEFT OUTER JOIN Clientes AS Fornecedores " & vbCrLf &
              "               ON ContasAPagar.Cliente = Fornecedores.Cliente_Id " & vbCrLf &
              "              AND ContasAPagar.EndCliente = Fornecedores.Endereco_Id" & vbCrLf &
              "            WHERE (ContasAPagar.RegistroMestre = " & titulo & ")) AS Consulta" & vbCrLf &
              "  ORDER BY Titulo" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "FilhosMestre")
    End Function

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            Relatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            Relatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Relatorio(ByVal pdf As Boolean)
        If Funcoes.VerificaPermissao("RelatorioDeTitulos", "RELATORIO") Then

            'ACRESCENTADO PARA CONSOLIDAR EMPRESA ONDE MELHORA A PERFORMANCE DO SQL - FURLAN - 16/10/2024
            If chkConsolidarEmpresa.Checked Then listarEmpresas()

            Dim Parametros As String = ""
            Dim crystal As String = ""
            Dim ds As DataSet

            If FinanceiroNovo Then
                ds = Banco.ConsultaDataSet(GetSQLTituloNovo(), "Titulos")
            Else
                ds = getDataSet(Parametros)
            End If

            If radData.Checked Then
                If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                    crystal = "Cr_Titulos"
                ElseIf FinanceiroNovo Then
                    crystal = "Cr_Titulos"
                Else
                    crystal = "Cr_TitulosPorEmpresa"
                End If
            ElseIf radClienteData.Checked Then
                crystal = "Cr_TitulosPorCliente"
            Else
                crystal = "Cr_TitulosPorPedido"
            End If

            Dim parameters = New Dictionary(Of String, Object)()

            If RdReceberPagar.Checked Then
                parameters.Add("Relatorio", "Relação de Títulos A Pagar e A Receber")
            Else
                parameters.Add("Relatorio", "Relação de Títulos A " & IIf(RdPagar.Checked, "Pagar", "Receber"))
            End If

            Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))

            parameters.Add("EmpresaNome", objEmpresa.Nome)
            parameters.Add("EmpresaCidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
            parameters.Add("EmpresaCodigo", Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

            parameters.Add("UnidadeDeNegocio", Parametros)
            parameters.Add("DataInicial", txtPeriodoInicialConsultaTitulos.Text)
            parameters.Add("DataFinal", txtPeriodoFinalConsultaTitulos.Text)
            parameters.Add("VisualizarProdutos", chkListarProdutos.Checked)
            parameters.Add("TipoDaCarteira", IIf(rdFinalidadeFinanceira.Checked, "Finalidade Financeira", "Carteira"))

            Funcoes.BindReport(Me.Page, ds, crystal, IIf(pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
        End If
    End Sub

    Private Function getDataSet(ByRef Parametros As String) As DataSet
        Dim ds As New DataSet

        Dim sql As String = String.Empty

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.Text) Then
            Parametros &= "Empresa: " & DdlEmpresaConsultaTitulos.SelectedItem.Text & vbCrLf
            If chkConsolidarEmpresa.Checked Then Parametros &= "Empresa Consolidada" & vbCrLf
        End If

        If txtCodigoCliente.Value.Length > 0 Then
            Parametros &= "Cliente: " & txtClientes.Text & vbCrLf
            If chkConsolidarCliente.Checked Then Parametros &= "Cliente Consolidado" & vbCrLf
        End If

        sql = "SELECT Empresas.Reduzido AS ReduzidoEmpresa, Titulos.Empresa, Titulos.EndEmpresa, Empresas.Nome AS NomeEmpresa, Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, " & vbCrLf &
              "        Titulos.Registro_Id as Titulo," & vbCrLf &
              "       convert(nvarchar,Titulos.Registro_Id) + ' ' + Case" & vbCrLf &
              "                                                       when isnull(Titulos.Moeda,1) = 1 " & vbCrLf &
              "                                                         then 'R$'" & vbCrLf &
              "                                                         Else 'U$'" & vbCrLf &
              "                                                     end Registro," & vbCrLf &
              "       '' as Faturamento," & vbCrLf &
              "       '' as Lote, " & vbCrLf &
              "       0 as LoteTotal," & vbCrLf &
              "       0 as LoteEntregue," & vbCrLf &
              "       Titulos.Pedido," & vbCrLf &
              "       Titulos.Cliente, Clientes.Nome AS NomeCliente," & vbCrLf &
              "       Titulos.Movimento, " & vbCrLf &
              "       Titulos.Prorrogacao AS Vencimento, " & vbCrLf &
              "       case when Titulos.Provisao = 1 then Titulos.Baixa else null end as Baixa, " & vbCrLf &
              "       Titulos.Carteira, " & vbCrLf &
              "       Carteira.Descricao AS NomeCarteira," & vbCrLf &
              "       Titulos.Historico" & IIf(chkObservacao.Checked, "+ ' / OBS: ' + cast(Titulos.Observacoes as varchar) as Historico,", ",") & vbCrLf &
              "       Titulos.solicitacao, " & vbCrLf &
              "       ISNULL(P.PedidoEfetivo,0) PedidoEfetivo, " & vbCrLf

        Dim Provisao As String


        Provisao = GetProvisao(radDataBase.Checked)

        sql &= Provisao & " As Provisao, " & vbCrLf

        '***************************************************************************
        '*************************** VALORES ***************************************
        '***************************************************************************

        sql &= "        case " & vbCrLf &
               "           when " & Provisao & " = 1 " & vbCrLf &
               "             then " & IIf(rbDocumento.Checked, "Titulos.ValorDoDocumento", "Titulos.ValorLiquido") & vbCrLf &
               "             else case " & vbCrLf &
               "                    when Titulos.Moeda = 1 " & vbCrLf &
               "                      then " & IIf(rbDocumento.Checked, "Titulos.ValorDoDocumento", "Titulos.ValorLiquido") & vbCrLf &
               "                      else convert(numeric(18,2), " & IIf(rbDocumento.Checked, "Titulos.MoedaValorDoDocumento", "Titulos.MoedaValorLiquido") & " * case when Titulos.Indexador = 99 then p.indiceFixado else Cotacoes.indice end) " & vbCrLf &
               "                   end " & vbCrLf

        'Soma as devolucoes
        'If radDataBase.Checked OrElse Excel Then

        If radDataBase.Checked Then
            sql &= "        end + ISNULL(Devolucoes.Valor,0) as ValorLiquido, " & vbCrLf
        Else
            sql &= "        end ValorLiquido, " & vbCrLf
        End If

        sql &= "        case " & vbCrLf &
               "           when " & Provisao & " = 1 " & vbCrLf &
               "			 then " & IIf(rbDocumento.Checked, "Titulos.MoedaValorDoDocumento", "Titulos.MoedaValorLiquido") & vbCrLf &
               "             else Case" & vbCrLf &
               "                    when Titulos.Moeda <> 1  " & vbCrLf &
               "                      then Titulos.MoedaValorLiquido" & vbCrLf &
               "                      else 0 " & vbCrLf &
               "                  end " & vbCrLf &
               "        end MoedaValorLiquido," & vbCrLf



        sql &= "      Titulos.Tipo AS tipo" & vbCrLf &
               IIf(Not radData.Checked, " into #temp", "") & vbCrLf

        sql &= GetFromWhereSql(False, Provisao, Parametros)

        If Not radData.Checked Then

            'Esta Consulta usa como critério a tabela #temp criada na consulta anteriormente onde já estão colocados todos os critérios necessários
            'Caso haja a necessidade de inclusão de novos critérios coloque-os apenas na consulta anterior.

            sql &= " SELECT * " & vbCrLf &
                   "   FROM #temp " & vbCrLf &
                   "  ORDER BY NomeCliente; " & vbCrLf &
                   "Select Titulo.Empresa, " & vbCrLf &
                   "       Titulo.EndEmpresa," & vbCrLf &
                   "       convert(nvarchar,Titulo.Registro) + ' ' +" & vbCrLf &
                   "       Case" & vbCrLf &
                   "         when isnull(Titulo.Moeda,1) = 1 " & vbCrLf &
                   "           then 'R$'" & vbCrLf &
                   "           Else 'U$'" & vbCrLf &
                   "       end Registro," & vbCrLf &
                   "       PxI.CodigoProduto," & vbCrLf &
                   "       PxI.NomeProduto," & vbCrLf &
                   "       PxI.Quantidade," & vbCrLf &
                   "       '' as Faturamento," & vbCrLf &
                   "       '' as Lote, " & vbCrLf &
                   "       0 as LoteTotal," & vbCrLf &
                   "       0 as LoteEntregue," & vbCrLf &
                   "       case " & vbCrLf &
                   "         when  PxI.Quantidade = 0" & vbCrLf &
                   "          then 0" & vbCrLf &
                   "          else PxI.Valor / PxI.Quantidade" & vbCrLf &
                   "       end as Unitario," & vbCrLf &
                   "       PxI.Valor," & vbCrLf &
                   "       PxI.UnidadeDeMedida" & vbCrLf &
                   " From (" & vbCrLf &
                   "	      Select Empresa, Endempresa, convert(nvarchar,Registro_id) as Registro, Pedido, Moeda" & vbCrLf &
                   "	        From ContasaPagar " & vbCrLf &
                   "	       Union all" & vbCrLf &
                   "	      Select Empresa, Endempresa, convert(nvarchar,Registro_id), Pedido, Moeda" & vbCrLf &
                   "	        From ContasaReceber" & vbCrLf &
                   "           Union all" & vbCrLf &
                   "	      Select Empresa_id, Endempresa_id, convert(nvarchar,Pedido_id) + '-'+  convert(nvarchar,parcela_id), Pedido_id, Moeda" & vbCrLf &
                   "	        From VW_TituloVirtual" & vbCrLf &
                   "	   ) Titulo" & vbCrLf &
                   " Inner Join (" & vbCrLf &
                   "				Select Pxi.Empresa_id," & vbCrLf &
                   "					   Pxi.EndEmpresa_id," & vbCrLf &
                   "					   Pxi.Pedido_Id," & vbCrLf &
                   "					   Pxi.Produto_Id as CodigoProduto," & vbCrLf &
                   "					   Prd.Nome as NomeProduto," & vbCrLf &
                   "					   Sum(case" & vbCrLf &
                   "						     When Pxi.TipoDeLancamento = 'E'" & vbCrLf &
                   "						  	   then Pxi.Quantidade * - 1" & vbCrLf &
                   "							   else Pxi.Quantidade" & vbCrLf &
                   "					       end) as Quantidade," & vbCrLf &
                   "					   Sum(case" & vbCrLf &
                   "						     When P.Moeda = 1 " & vbCrLf &
                   "						  	   then case " & vbCrLf &
                   "                                   When Pxi.TipoDeLancamento = 'E'" & vbCrLf &
                   "                                     then Pxi.TotalOficial * -1 " & vbCrLf &
                   "                                     else Pxi.TotalOficial " & vbCrLf &
                   "                                 end" & vbCrLf &
                   "							   else case " & vbCrLf &
                   "                                   When Pxi.TipoDeLancamento = 'E'" & vbCrLf &
                   "                                     then Pxi.TotalMoeda * -1 " & vbCrLf &
                   "                                     else Pxi.TotalMoeda " & vbCrLf &
                   "                                 end" & vbCrLf &
                   "					       end) as Valor," & vbCrLf &
                   "					   Prd.Unidade as UnidadeDeMedida" & vbCrLf &
                   "				  from Pedidos P" & vbCrLf &
                   "              INNER JOIN PedidoXItemXLancamento Pxi" & vbCrLf &
                   "                 ON P.Empresa_Id    = Pxi.Empresa_Id " & vbCrLf &
                   "                AND P.EndEmpresa_Id = Pxi.EndEmpresa_Id " & vbCrLf &
                   "                AND P.Pedido_Id     = Pxi.Pedido_Id " & vbCrLf &
                   "				 Inner join Produtos Prd" & vbCrLf &
                   "				    on Pxi.Produto_Id = Prd.Produto_Id" & vbCrLf &
                   "              WHERE P.Pedido_Id IN (SELECT Pedido from #temp) " & vbCrLf &
                   "              group by Pxi.Empresa_id, Pxi.EndEmpresa_id, Pxi.Pedido_Id, Pxi.Produto_Id, Prd.Nome, Prd.Unidade" & vbCrLf &
                   "             ) PxI" & vbCrLf &
                   "   on PxI.Empresa_id    = Titulo.Empresa" & vbCrLf &
                   "  and PxI.EndEmpresa_id = Titulo.EndEmpresa" & vbCrLf &
                   "  and PxI.Pedido_id     = Titulo.Pedido" & vbCrLf &
                   " where Titulo.Registro in (Select Titulo from #Temp)" & vbCrLf
            ds = Banco.ConsultaDataSet(sql, "Titulos")
            ds.Tables(0).TableName = "Titulos"
            ds.Tables(1).TableName = "PedidoProduto"
        Else
            ds = Banco.ConsultaDataSet(sql, "Titulos")
        End If
        Return ds
    End Function

    Private Function getDataSetExcel() As DataSet
        Dim Parametros As String = String.Empty
        Dim Sql As String = String.Empty

        Sql = "SELECT Titulos.Registro_Id AS Titulo, " & vbCrLf

        If RdReceber.Checked Then
            Sql &= "       'R-' + Case" & vbCrLf
        ElseIf RdPagar.Checked Then
            Sql &= "       'P-' + Case" & vbCrLf
        Else
            Sql &= "       '(RP)-' + Case" & vbCrLf
        End If

        Sql &= "                when Titulos.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "' and Titulos.provisao = 1 " & vbCrLf &
        "                  then '2' " & vbCrLf &
        "                  Else convert(varchar, Titulos.provisao) " & vbCrLf &
        "              end as Provisao, " & vbCrLf &
        "       Titulos.ContaContabilPagadora, " & vbCrLf &
        "       ISNULL(PLCC.Titulo,'') AS NomeDaConta," & vbCrLf &
        "       Carteiras.Produto_Id AS CarteiraFinanceira," & vbCrLf &
        "       Carteiras.Descricao AS Descricaocf," & vbCrLf &
        "       Titulos.CarteiraDoTitulo AS CarteiraDoTitulo," & vbCrLf &
        "       Carteira.Descricao AS Descricaoct," & vbCrLf &
        "       Titulos.Tributo AS Encargo," & vbCrLf &
        "       ISNULL(Encargos.Descricao,'') AS NomeEncargo," & vbCrLf &
        "       Titulos.Moeda," & vbCrLf &
        "       CONVERT(VARCHAR,Titulos.Movimento,103) AS Movimento," & vbCrLf &
        "       CONVERT(VARCHAR,Titulos.Vencimento,103) AS Vencimento," & vbCrLf &
        "       CONVERT(VARCHAR,Titulos.Prorrogacao,103) AS Prorrogacao," & vbCrLf &
        "       case" & vbCrLf &
        "         when Titulos.provisao = 1" & vbCrLf &
        "            then CONVERT(VARCHAR,Titulos.Baixa,103)" & vbCrLf &
        "            else ''" & vbCrLf &
        "         end AS Baixa," & vbCrLf &
        "       Emp.Reduzido AS Empresa," & vbCrLf &
        "       Emp.Nome AS NomeDaEmpresa," & vbCrLf &
        "       Emp.Cidade AS CidadeDaEmpresa," & vbCrLf &
        "       GP.Conta_Id AS GrupodeContas," & vbCrLf &
        "       GP.Titulo AS NomeGrupoDeContas," & vbCrLf &
        "       Titulos.Cliente AS Fornecedor," & vbCrLf &
        "       Clientes.Nome AS NomeDoFornecedor," & vbCrLf &
        "       ISNULL(PCE.Conta_Id,'') AS ContaContabilEncargo," & vbCrLf &
        "       ISNULL(PCE.Titulo,'') AS NomeContaEncargo," & vbCrLf &
        "       Titulos.Pedido," & vbCrLf

        If radDataBase.Checked Then
            Sql &= "         Titulos.ValorDoDocumento + ISNULL(Devolucoes.Valor,0) as ValorDoDocumento, " & vbCrLf
        Else
            Sql &= "         Titulos.ValorDoDocumento, " & vbCrLf
        End If

        Sql &= "       Titulos.Descontos," & vbCrLf &
                  "       Titulos.Deducoes," & vbCrLf &
                  "       Titulos.Juros," & vbCrLf &
                  "       Titulos.Acrescimos," & vbCrLf

        If radDataBase.Checked Then
            Sql &= "        Titulos.ValorLiquido + ISNULL(Devolucoes.Valor,0) as ValorLiquido, " & vbCrLf
        Else
            Sql &= "        Titulos.ValorLiquido, " & vbCrLf
        End If

        Sql &= "          Titulos.Historico," & vbCrLf &
                  "       Titulos.Observacoes," & vbCrLf &
                  "       ISNULL(Dest.Cliente_Id,'') AS Destinacao," & vbCrLf &
                  "       ISNULL(Dest.Nome,'') AS NomeDoDestinatario," & vbCrLf &
                  "       ISNULL((SELECT top 1 Produtos.Produto_Id " & vbCrLf &
                  "                 FROM PedidoXItem AS pXi" & vbCrLf &
                  "                INNER JOIN Produtos " & vbCrLf &
                  "                   ON pXi.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                  "                  AND Produtos.Agrupar = 'N'" & vbCrLf &
                  "                WHERE pXi.Empresa_Id    = Titulos.Empresa" & vbCrLf &
                  "                  AND pXi.EndEmpresa_Id = Titulos.EndEmpresa" & vbCrLf &
                  "                  AND pXi.Pedido_Id     = Titulos.Pedido),'') AS Produto," & vbCrLf &
                  "       ISNULL((SELECT top 1 Produtos.Nome" & vbCrLf &
                  "                 FROM PedidoXItem AS pXi" & vbCrLf &
                  "                INNER JOIN Produtos " & vbCrLf &
                  "  	                ON pXi.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                  "                  AND Produtos.Agrupar = 'N'" & vbCrLf &
                  "                WHERE pXi.Empresa_Id    = Titulos.Empresa" & vbCrLf &
                  "                  AND pXi.EndEmpresa_Id = Titulos.EndEmpresa" & vbCrLf &
                  "                  AND pXi.Pedido_Id     = Titulos.Pedido),'') AS NomeProduto," & vbCrLf

        Sql &= "   (SELECT SUM(EncargoLiquido.Valor) FROM  NotasFiscaisXItens AS ItensNota     " & vbCrLf
        Sql &= "    INNER Join NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNota.Empresa_Id    " & vbCrLf
        Sql &= "    And EncargoLiquido.EndEmpresa_Id = ItensNota.EndEmpresa_Id     " & vbCrLf
        Sql &= "    And EncargoLiquido.Cliente_Id = ItensNota.Cliente_Id     " & vbCrLf
        Sql &= "    And EncargoLiquido.EndCliente_Id = ItensNota.EndCliente_Id     " & vbCrLf
        Sql &= "    And EncargoLiquido.EntradaSaida_Id = ItensNota.EntradaSaida_Id    " & vbCrLf
        Sql &= "    And EncargoLiquido.Serie_Id = ItensNota.Serie_Id   " & vbCrLf
        Sql &= "    And EncargoLiquido.Nota_Id = ItensNota.Nota_Id    " & vbCrLf
        Sql &= "    And EncargoLiquido.Produto_Id = ItensNota.Produto_Id     " & vbCrLf
        Sql &= "    And EncargoLiquido.Encargo_Id = 'LIQUIDO'     " & vbCrLf
        Sql &= "    WHERE ItensNota.Empresa_Id = nf.Empresa_Id     " & vbCrLf
        Sql &= "    And ItensNota.EndEmpresa_Id = nf.EndEmpresa_Id     " & vbCrLf
        Sql &= "    And ItensNota.Cliente_Id = nf.Cliente_Id    " & vbCrLf
        Sql &= "    And ItensNota.EndCliente_Id = nf.EndCliente_Id     " & vbCrLf
        Sql &= "    And ItensNota.EntradaSaida_Id = nf.EntradaSaida_Id    " & vbCrLf
        Sql &= "    And ItensNota.Serie_Id = nf.Serie_Id    " & vbCrLf
        Sql &= "    And ItensNota.Nota_Id = nf.Nota_Id     " & vbCrLf
        Sql &= "   ) AS Total,     " & vbCrLf


        Sql &= "(SELECT ISNULL(SUM(CASE" & vbCrLf &
                  "                            WHEN TipoDeLancamento = 'E'" & vbCrLf &
                  "                              THEN ISNULL(pXi.TotalMoeda, 0) * - 1" & vbCrLf &
                  "                              ELSE ISNULL(pXi.TotalMoeda, 0)" & vbCrLf &
                  "                          END),0)" & vbCrLf &
                  "          FROM Pedidos" & vbCrLf &
                  "          LEFT JOIN PedidoXItemxLancamento AS pXi" & vbCrLf &
                  "            ON Pedidos.Pedido_Id     = pXi.Pedido_Id" & vbCrLf &
                  "           AND Pedidos.Empresa_Id    = pXi.Empresa_Id" & vbCrLf &
                  "           AND Pedidos.EndEmpresa_Id = pXi.EndEmpresa_Id	" & vbCrLf &
                  "         WHERE Pedidos.Pedido_Id     = Titulos.Pedido" & vbCrLf &
                  "           AND Pedidos.Empresa_Id    = Titulos.Empresa" & vbCrLf &
                  "           AND Pedidos.EndEmpresa_Id = Titulos.EndEmpresa) AS TotalEmDolar," & vbCrLf &
                  "       ISNULL(nfXt.Nota_Id,'') AS Nota," & vbCrLf &
                  "       ISNULL(CONVERT(VARCHAR,nf.Movimento,103),'') AS MovimentoNota," & vbCrLf &
                  "       ISNULL(Titulos.ContratoBancario,'') AS ContratoBanco," & vbCrLf &
                  "       ISNULL(Titulos.RegistroMestre,'') AS RegistroMestre," & vbCrLf &
                  "       ISNULL(sub.Classe,'') AS Classe," & vbCrLf &
                  "       case" & vbCrLf &
                  "          when len(isnull(Titulos.Tributo,'')) > 0" & vbCrLf &
                  "             then Encargos.ContaDebito" & vbCrLf &
                  "             else Carteiras.ContaClientes" & vbCrLf &
                  "             end AS Contrapartida," & vbCrLf &
                  "       PCE.Titulo as DescricaoContrapartida," & vbCrLf &
                  "       case" & vbCrLf &
                  "           when isnull(PCxBC.Adiantamento,0) = 1" & vbCrLf &
                  "               then 'C'" & vbCrLf &
                  "               else 'N'" & vbCrLf &
                  "           end AS TipoDaConta," & vbCrLf &
                  "	   isnull((SELECT top 1 raz.Conta_id " & vbCrLf &
                  "	      from razao raz " & vbCrLf &
                  "			where raz.empresa_id    = nf.Empresa_Id " & vbCrLf &
                  "			and raz.endempresa_id   = nf.EndEmpresa_Id " & vbCrLf &
                  "			and raz.cliente_nf      = nf.Cliente_Id " & vbCrLf &
                  "			and raz.endcliente_nf   = nf.EndCliente_Id " & vbCrLf &
                  "			and raz.entradasaida_nf = nf.EntradaSaida_Id " & vbCrLf &
                  "			and raz.serie_nf        = nf.Serie_Id " & vbCrLf &
                  "			and raz.numero_nf       = nf.Nota_Id " & vbCrLf &
                  "			and raz.Encargo_NF = 'PRODUTO'),'') AS ContrapartidaNota," & vbCrLf &
                  "	   isnull((SELECT top 1 raz.Conta_id as Titulo " & vbCrLf &
                  "	      from razao raz " & vbCrLf &
                  "			where raz.empresa_id    = nf.Empresa_Id " & vbCrLf &
                  "			and raz.endempresa_id   = nf.EndEmpresa_Id " & vbCrLf &
                  "			and raz.cliente_nf      = nf.Cliente_Id" & vbCrLf &
                  "			and raz.endcliente_nf   = nf.EndCliente_Id " & vbCrLf &
                  "			and raz.entradasaida_nf = nf.EntradaSaida_Id " & vbCrLf &
                  "			and raz.serie_nf        = nf.Serie_Id " & vbCrLf &
                  "			and raz.numero_nf       = nf.Nota_Id " & vbCrLf &
                  "			and raz.Encargo_NF = 'PRODUTO'),'') AS DescricaoContrapartidaNota, Tipo " & vbCrLf


        Dim Provisao As String = GetProvisao(radDataBase.Checked)
        Sql &= GetFromWhereSql(True, Provisao, Parametros)

        Return Banco.ConsultaDataSet(Sql, "Titulos")

    End Function

    Private Function GetFromWhereSql(ByVal Excel As Boolean, ByVal Provisao As String, Optional ByRef parametros As String = "") As String
        Dim SqlT As String = String.Empty

        '***************************************************************************
        '********** CONTAS A PAGAR / RECEBER ***************************************
        '***************************************************************************
        SqlT &= " FROM ("

        If RdReceberPagar.Checked Or RdPagar.Checked Then
            SqlT &= "       SELECT 'P' as Tipo, Empresa, Endempresa, convert(nvarchar,Registro_id) as Registro_id, EmpresaPedido, EndEmpresaPedido, Pedido, Moeda, Prorrogacao, Indexador, Cliente, EndCliente, Carteira, " & vbCrLf &
                    "              CarteiraDoTitulo, situacao, Grupado, ContacontabilCliente, Provisao, MoedaValorDoDocumento, UnidadeDeNegocio, Baixa, Movimento," & vbCrLf &
                    "              Historico, Solicitacao, valorliquido, ValorDoDocumento, MoedaValorLiquido, UsuarioInclusaoData, PedidoFixacao, Observacoes," & vbCrLf &
                    "              ContaContabilPagadora, Tributo, Vencimento, Juros, Acrescimos, Descontos, Deducoes, RegistroMestre, ContratoBancario, " & vbCrLf &
                    "              Destinatario, EndDestinatario, TipoPagto, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora " & vbCrLf &
                    "         FROM ContasAPagar" & vbCrLf
            If RdReceberPagar.Checked Then SqlT &= " UNION ALL"
        End If

        If RdReceberPagar.Checked Or RdReceber.Checked Then
            SqlT &= "       SELECT 'R' AS Tipo, Empresa, Endempresa, convert(nvarchar,Registro_id) as Registro_id, EmpresaPedido, EndEmpresaPedido, Pedido, Moeda, Prorrogacao, Indexador, Cliente, EndCliente, Carteira, " & vbCrLf &
                    "              CarteiraDoTitulo, situacao, Grupado, ContacontabilCliente, Provisao, MoedaValorDoDocumento, UnidadeDeNegocio, Baixa, Movimento," & vbCrLf &
                    "              Historico, Solicitacao, valorliquido, ValorDoDocumento, MoedaValorLiquido, UsuarioInclusaoData, PedidoFixacao, Observacoes," & vbCrLf &
                    "              ContaContabilPagadora, Tributo, Vencimento, Juros, Acrescimos, Descontos, Deducoes, RegistroMestre, ContratoBancario, " & vbCrLf &
                    "              Destinatario, EndDestinatario, TipoPagto, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora " & vbCrLf &
                    "         FROM ContasaReceber" & vbCrLf
        End If

        If chkProjecao.Checked Then
            SqlT &= "       Union ALL" & vbCrLf &
                    "       Select V.CR, V.Empresa_id, V.EndEmpresa_Id, convert(nvarchar,V.Pedido_id) + '-' + Convert(nvarchar,V.Parcela_id), V.Empresa_id, V.EndEmpresa_id, V.Pedido_id, V.moeda,V.Vencimento, 3 ,V.cliente,V.endcliente,null," & vbCrLf &
                    "              null,1,'N',null,2,v.ParcelaMoeda,p.UnidadeDeNegocio,Vencimento,Vencimento," & vbCrLf &
                    "       	   'Parcela Virtual do Pedido ' + convert(nvarchar,V.Pedido_id) + '-' + Convert(nvarchar,V.Parcela_id), 0,v.ParcelaOficial,v.ParcelaOficial,v.ParcelaMoeda,null,null,null," & vbCrLf &
                    "       	   null,null,v.Vencimento,0,0,0,0,0,0," & vbCrLf &
                    "       	   null,null,null, null, null, null, null, null, null, null " & vbCrLf &
                    "         from VW_TituloVirtual V" & vbCrLf &
                    "        inner join Pedidos P" & vbCrLf &
                    "           on v.Empresa_id    = p.Empresa_Id" & vbCrLf &
                    "          and v.EndEmpresa_id = p.EndEmpresa_Id" & vbCrLf &
                    "          and v.Pedido_id     = p.Pedido_Id" & vbCrLf

            If RdReceber.Checked Then
                SqlT &= "  Where V.cr = 'R'"
            ElseIf RdPagar.Checked Then
                SqlT &= "  Where V.cr = 'P'"
            End If
        End If

        SqlT &= "       )AS Titulos" & vbCrLf

        '***************************************************************************
        '***************************************************************************

        SqlT &= "  LEFT Join cotacoes" & vbCrLf &
                "     on Cotacoes.Data_id      = Titulos.Prorrogacao " & vbCrLf &
                "    and Cotacoes.Indexador_Id = Titulos.Indexador " & vbCrLf &
                "  INNER JOIN Clientes AS Empresas " & vbCrLf &
                "     ON Titulos.Empresa    = Empresas.Cliente_Id " & vbCrLf &
                "    AND Titulos.EndEmpresa = Empresas.Endereco_Id " & vbCrLf &
                "  LEFT JOIN Clientes AS EmpresasPagadora " & vbCrLf &
                "     ON Titulos.EmpresaPagadora    = EmpresasPagadora.Cliente_Id " & vbCrLf &
                "    AND Titulos.EndEmpresaPagadora = EmpresasPagadora.Endereco_Id " & vbCrLf &
                "  INNER JOIN Clientes " & vbCrLf &
                "     ON Titulos.Cliente    = Clientes.Cliente_Id " & vbCrLf &
                "    AND Titulos.EndCliente = Clientes.Endereco_Id " & vbCrLf &
                "   LEFT OUTER JOIN ComprasXProdutos AS Carteiras " & vbCrLf &
                "     ON Titulos.Carteira = Carteiras.Produto_Id" & vbCrLf &
                "   LEFT OUTER JOIN Carteira " & vbCrLf &
                "     ON Titulos.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf

        'If radDataBase.Checked OrElse Excel Then
        '    SqlT &= "   inner join NotaFiscalXTitulo nfXt " & vbCrLf & _
        '            "     on Titulos.Registro_Id = NFxT.Titulo_Id " & vbCrLf & _
        '            "    and Titulos.Empresa     = NFxT.Empresa_Id " & vbCrLf & _
        '            "    and Titulos.EndEmpresa  = NFxT.EndEmpresa_Id " & vbCrLf & _
        '            "    and Titulos.Cliente     = NFxT.Cliente_Id " & vbCrLf & _
        '            "    and Titulos.EndCliente  = NFxT.EndCliente_Id " & vbCrLf & _
        '            "   inner join NotasFiscais nf " & vbCrLf & _
        '            "     on nf.Nota_Id         = nfXt.Nota_Id " & vbCrLf & _
        '            "    and nf.Serie_Id        = nfXt.Serie_Id " & vbCrLf & _
        '            "    and nf.EntradaSaida_Id = nfXt.EntradaSaida_Id " & vbCrLf & _
        '            "    and nf.Empresa_Id      = nfXt.Empresa_Id " & vbCrLf & _
        '            "    and nf.EndEmpresa_Id   = nfXt.EndEmpresa_Id " & vbCrLf & _
        '            "    and nf.Cliente_Id      = nfXt.Cliente_Id " & vbCrLf & _
        '            "    and nf.EndCliente_Id   = nfXt.EndCliente_Id " & vbCrLf & _
        '            "    and nf.movimento      <='" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"


        'End If

        If radDataBase.Checked OrElse Excel Then

            SqlT &= "   left join NotaFiscalXTitulo nfXt " & vbCrLf &
                    "     on Titulos.Registro_Id = NFxT.Titulo_Id " & vbCrLf &
                    "    and Titulos.Empresa     = NFxT.Empresa_Id " & vbCrLf &
                    "    and Titulos.EndEmpresa  = NFxT.EndEmpresa_Id " & vbCrLf &
                    "    and Titulos.Cliente     = NFxT.Cliente_Id " & vbCrLf &
                    "    and Titulos.EndCliente  = NFxT.EndCliente_Id " & vbCrLf &
                    "   left join NotasFiscais nf " & vbCrLf &
                    "     on nf.Nota_Id         = nfXt.Nota_Id " & vbCrLf &
                    "    and nf.Serie_Id        = nfXt.Serie_Id " & vbCrLf &
                    "    and nf.EntradaSaida_Id = nfXt.EntradaSaida_Id " & vbCrLf &
                    "    and nf.Empresa_Id      = nfXt.Empresa_Id " & vbCrLf &
                    "    and nf.EndEmpresa_Id   = nfXt.EndEmpresa_Id " & vbCrLf &
                    "    and nf.Cliente_Id      = nfXt.Cliente_Id " & vbCrLf &
                    "    and nf.EndCliente_Id   = nfXt.EndCliente_Id " & vbCrLf &
                    "    and nf.movimento      <='" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"

            'Rodrigo 03/05/2023 - Adicionado apuracao de valor liquido da nota e saldo de titulos baixados para verificacao de saldo fechado'
            SqlT &= "   OUTER APPLY(SELECT SUM(EncargoLiquidoNota.Valor) as Total " & vbCrLf &
                    "               FROM NotasFiscaisXEncargos As EncargoLiquidoNota " & vbCrLf &
                    "               WHERE EncargoLiquidoNota.Empresa_Id = nf.Empresa_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.Cliente_Id = nf.Cliente_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.EndCliente_Id = nf.EndCliente_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.Serie_Id = nf.Serie_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.Nota_Id = nf.Nota_Id " & vbCrLf &
                    "               AND   EncargoLiquidoNota.Encargo_Id = 'LIQUIDO') AS  EncargoLiquido"

            SqlT &= "   OUTER APPLY(SELECT SUM(Contas.ValorDoDocumento) as Total " & vbCrLf &
                    "               FROM NotaFiscalXTitulo As Titulo " & vbCrLf


            If RdReceber.Checked Then
                SqlT &= " INNER Join ContasAReceber AS Contas ON Contas.Registro_Id = Titulo.Titulo_Id " & vbCrLf
            Else
                SqlT &= "INNER Join ContasAPagar AS Contas ON Contas.Registro_Id = Titulo.Titulo_Id " & vbCrLf
            End If

            SqlT &= " WHERE Titulo.Empresa_Id = nf.Empresa_Id " & vbCrLf &
                    " AND   Titulo.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf &
            "         AND   Titulo.Cliente_Id = nf.Cliente_Id" & vbCrLf &
            "         AND   Titulo.EndCliente_Id = nf.EndCliente_Id " & vbCrLf &
            "         AND   Titulo.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
            "         AND   Titulo.Serie_Id = nf.Serie_Id " & vbCrLf &
            "         AND   Titulo.Nota_Id = nf.Nota_Id " & vbCrLf &
            "         AND   Contas.Provisao = 1 ) TitulosBaixados" & vbCrLf


            'Rodrigo 26/08/2022 - Divide pelo total de títulos para resolver problemas quando tiver mais de um título
            SqlT &= " OUTER APPLY (SELECT SUM(Devolucoes.ValorLiquido) / (SELECT COUNT(Titulo_Id) as Contador  " & vbCrLf
            SqlT &= "                                                  FROM NotaFiscalXTitulo AS ContadorDeTitulo " & vbCrLf
            SqlT &= "                                                  WHERE ContadorDeTitulo.Empresa_Id = NF.Empresa_Id " & vbCrLf
            SqlT &= "                                                  AND   ContadorDeTitulo.EndEmpresa_Id = NF.EndEmpresa_Id " & vbCrLf
            SqlT &= "                                                  AND   ContadorDeTitulo.Cliente_Id = NF.Cliente_Id " & vbCrLf
            SqlT &= "                                                  AND   ContadorDeTitulo.EndCliente_Id = NF.EndCliente_Id " & vbCrLf
            SqlT &= "                                                  AND   ContadorDeTitulo.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf
            SqlT &= "                                                  AND   ContadorDeTitulo.Serie_Id = NF.Serie_Id" & vbCrLf
            SqlT &= "                                                  AND   ContadorDeTitulo.Nota_Id = NF.Nota_Id) AS Valor " & vbCrLf & vbCrLf

            SqlT &= "              FROM NotaFiscalDevolucaoXNotaFiscal as Devolucoes " & vbCrLf & vbCrLf

            SqlT &= "              INNER JOIN NotasFiscais as NotaDevolucao on NotaDevolucao.Empresa_Id = Devolucoes.EmpresaDevolucao_Id " & vbCrLf
            SqlT &= "              AND NotaDevolucao.EndEmpresa_Id = Devolucoes.EndEmpresaDevolucao_Id  " & vbCrLf
            SqlT &= "              AND NotaDevolucao.Cliente_Id = Devolucoes.ClienteDevolucao_Id " & vbCrLf
            SqlT &= "              AND NotaDevolucao.EndCliente_Id = Devolucoes.EndClienteDevolucao_Id " & vbCrLf
            SqlT &= "              AND NotaDevolucao.EntradaSaida_Id = Devolucoes.EntradaSaidaDevolucao_Id " & vbCrLf
            SqlT &= "              AND NotaDevolucao.Serie_Id = Devolucoes.SerieDevolucao_Id " & vbCrLf
            SqlT &= "              AND NotaDevolucao.Nota_Id = Devolucoes.NotaDevolucao_Id " & vbCrLf
            SqlT &= "              AND NotaDevolucao.Movimento > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf & vbCrLf

            SqlT &= "              INNER JOIN NotasFiscaisXItens as ItensNotaDevolucao on ItensNotaDevolucao.Empresa_Id = NotaDevolucao.Empresa_Id " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.EndEmpresa_Id = NotaDevolucao.EndEmpresa_Id   " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.Cliente_Id = NotaDevolucao.Cliente_Id " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.EndCliente_Id = NotaDevolucao.EndCliente_Id " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.EntradaSaida_Id = NotaDevolucao.EntradaSaida_Id " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.Serie_Id = NotaDevolucao.Serie_Id " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.Nota_Id = NotaDevolucao.Nota_Id " & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.CFOP_Id = Devolucoes.CFOPDevolucao_Id " & vbCrLf & vbCrLf
            SqlT &= "              AND ItensNotaDevolucao.Produto_Id = Devolucoes.Produto_Id "

            SqlT &= "              INNER JOIN NotasFiscaisXEncargos as EncargoLiquido on EncargoLiquido.Empresa_Id = ItensNotaDevolucao.Empresa_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id  " & vbCrLf
            SqlT &= "              AND EncargoLiquido.Cliente_Id = ItensNotaDevolucao.Cliente_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.Serie_Id = ItensNotaDevolucao.Serie_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.Nota_Id = ItensNotaDevolucao.Nota_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.Produto_Id = ItensNotaDevolucao.Produto_Id " & vbCrLf
            SqlT &= "              AND EncargoLiquido.Encargo_Id = 'LIQUIDO' " & vbCrLf & vbCrLf

            SqlT &= "WHERE Devolucoes.EmpresaDevolucao_Id = nf.Empresa_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.EndEmpresaDevolucao_Id = nf.EndEmpresa_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.ClienteDevolucao_Id  = nf.Cliente_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.EndClienteDevolucao_Id = nf.EndCliente_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.Produto_Id  = ItensNotaDevolucao.Produto_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.Serie_Id = nf.Serie_Id " & vbCrLf
            SqlT &= "AND   Devolucoes.Nota_Id = nf.Nota_Id " & vbCrLf & vbCrLf

            SqlT &= " ) as Devolucoes " & vbCrLf

        End If

        SqlT &= "   Left Join Pedidos AS P " & vbCrLf &
               "     ON P.Empresa_id    = Titulos.EmpresaPedido" & vbCrLf &
               "    And P.EndEmpresa_Id = Titulos.EndEmpresaPedido " & vbCrLf &
               "    And P.Pedido_id     = Titulos.Pedido " & vbCrLf

        If ChkApenasNF.Checked Then
            SqlT &= "   Inner Join NotaFiscalXTitulo AS NxT " & vbCrLf &
                   "     ON NxT.Empresa_id    = Titulos.EmpresaPedido" & vbCrLf &
                   "    And NxT.EndEmpresa_Id = Titulos.EndEmpresaPedido " & vbCrLf &
                   "    And NxT.Titulo_Id     = Titulos.Registro_Id " & vbCrLf
        End If

        '/*Utilizado caso seja gerado Excel Dados*/
        If Excel Then
            SqlT &= "   LEFT JOIN PlanoDeContas AS PLCC" & vbCrLf &
                    "     ON Titulos.ContaContabilPagadora = PLCC.Conta_Id" & vbCrLf &
                    "   LEFT JOIN clientes AS Emp" & vbCrLf &
                    "     ON Titulos.Empresa = Emp.Cliente_Id" & vbCrLf &
                    "    AND Titulos.EndEmpresa = Emp.Endereco_Id" & vbCrLf &
                    "   LEFT JOIN PlanoDeContas AS GP" & vbCrLf &
                    "     ON Titulos.ContaContabilCliente = GP.Conta_Id" & vbCrLf &
                    "   LEFT JOIN Encargos" & vbCrLf &
                    "     ON Titulos.Tributo = Encargos.Encargo_id" & vbCrLf &
                    "   LEFT JOIN PlanoDeContas AS PCE" & vbCrLf &
                    "     ON PCE.Conta_Id = case when len(isnull(Titulos.Tributo,'')) = 0 then Carteiras.ContaClientes else Encargos.ContaDebito end " & vbCrLf &
                    "  LEFT JOIN Clientes AS Dest" & vbCrLf &
                    "    ON Titulos.Destinatario    = Dest.Cliente_Id" & vbCrLf &
                    "   AND Titulos.EndDestinatario = Dest.Endereco_Id" & vbCrLf &
                    "  LEFT JOIN SubOperacoes AS sub" & vbCrLf &
                    "    ON P.Operacao    = sub.Operacao_Id" & vbCrLf &
                    "   AND P.SubOperacao = sub.SubOperacoes_Id" & vbCrLf &
                    "  LEFT JOIN BancosXContas bXc " & vbCrLf &
                    "	ON bXc.Empresa_Id        = Titulos.EmpresaPagadora " & vbCrLf &
                    "	AND bXc.EndEmpresa_Id    = Titulos.EndEmpresaPagadora " & vbCrLf &
                    "	AND bXc.Banco_Id         = Titulos.BancoPagador " & vbCrLf &
                    "	AND bXc.Agencia_Id       = Titulos.AgenciaPagadora " & vbCrLf &
                    "	AND bXc.DigitoAgencia_Id = Titulos.DigitoAgenciaPagadora " & vbCrLf &
                    "	AND bXc.Conta_Id         = Titulos.ContaPagadora " & vbCrLf &
                    "	AND bXc.DigitoConta_Id   = Titulos.DigitoContaPagadora " & vbCrLf &
                    "   LEFT JOIN PlanoDeContas AS PCxBC " & vbCrLf &
                    "     ON PCxBC.Conta_Id = bXc.ContaContabil " & vbCrLf
        End If

        If ddlSafra.SelectedIndex > 0 Then
            SqlT &= "  Where P.Safra = '" & ddlSafra.SelectedValue & "' " & vbCrLf
            parametros &= "safra: " & ddlSafra.SelectedValue & vbCrLf
        Else
            SqlT &= "  Where 1 = 1"
        End If

        Dim op As String = String.Join("','", lstFinalidade.GetSelecteds)
        If op.Length > 0 Then
            SqlT &= "   AND isnull(P.Finalidade,0) " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & op & "')"
            parametros &= "Finalidades do Pedido:" & String.Join(",", lstFinalidade.GetSelectedItems)
        End If


        Dim TipoPagto As String = String.Join("','", lstTipoPagRec.GetSelecteds)
        If TipoPagto.Length > 0 Then
            SqlT &= "              AND isnull(Titulos.TipoPagto,0) " & IIf(rdComTpPagRec.Checked, "", "not") & " in ('" & TipoPagto & "')" & vbCrLf
        End If

        If ddlRepresentante.SelectedIndex > 0 Then
            SqlT &= "  And exists (Select 1" & vbCrLf &
                    "                from Comissoes C" & vbCrLf &
                    "               where C.Empresa_id          = P.Empresa_id" & vbCrLf &
                    "                 And C.EndEmpresa_Id       = P.EndEmpresa_Id " & vbCrLf &
                    "                 And C.Pedido_id           = P.Pedido_id " & vbCrLf &
                    "                 And C.Representante_id    ='" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                    "                 And C.EndRepresentante_Id = " & ddlRepresentante.SelectedValue.Split("-")(1) & ")" & vbCrLf

            parametros &= "Representante: " & ddlRepresentante.SelectedItem.Text & vbCrLf
        End If

        If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
            SqlT &= "   AND Titulos.Situacao in(1,101,102,105) " & vbCrLf
        Else
            SqlT &= "   AND Titulos.Situacao in(1,101,102) " & vbCrLf
        End If

        SqlT &= "   AND (Titulos.Grupado <> 'M') " & vbCrLf

        'Comentei por não estar aparecendo os Titulos Lançados Tirando do Banco para o Caixa e de Transferência entre Contas - 31/03/2016 - Furlan
        '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf & _
        '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf

        If radDataBase.Checked Then


            '/* comentado 2024-10-08 devido a devolução do titulo 9341 RT */

            ''Rodrigo 03/05/2023 - Se tem algum título aberto ou já foi liquidado - so n vai vim os que tem saldo aberto mas n tem titulo, isso será tratado no terceiro union'
            'SqlT &= "   And (EXISTS " & vbCrLf &
            '        "         (SELECT TOP 1 * " & vbCrLf &
            '        "          FROM NotaFiscalXTitulo As Titulo " & vbCrLf


            'If RdReceber.Checked Then
            '    SqlT &= " INNER Join ContasAReceber ON ContasAReceber.Registro_Id = Titulo.Titulo_Id " & vbCrLf

            '    SqlT &= " WHERE Titulo.Empresa_Id = nf.Empresa_Id " & vbCrLf &
            '            " AND   Titulo.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf &
            '    "         AND   Titulo.Cliente_Id = nf.Cliente_Id" & vbCrLf &
            '    "         AND   Titulo.EndCliente_Id = nf.EndCliente_Id " & vbCrLf &
            '    "         AND   Titulo.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
            '    "         AND   Titulo.Serie_Id = nf.Serie_Id " & vbCrLf &
            '    "         AND   Titulo.Nota_Id = nf.Nota_Id " & vbCrLf &
            '    "         AND   ContasAReceber.Provisao = 2" & vbCrLf &
            '    "         And CASE " & vbCrLf &
            '    "               WHEN ContasAReceber.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
            '    "               	 And ContasAReceber.Provisao = 1 THEN 2 " & vbCrLf &
            '    "               ELSE ContasAReceber.Provisao" & vbCrLf &
            '    "             END in(2,3))" & vbCrLf &
            '    " Or (EncargoLiquido.Total =  TitulosBaixados.Total))" & vbCrLf

            'Else
            '    SqlT &= " INNER Join ContasAPagar ON ContasAPagar.Registro_Id = Titulo.Titulo_Id " & vbCrLf

            '    SqlT &= " WHERE Titulo.Empresa_Id = nf.Empresa_Id " & vbCrLf &
            '            " AND   Titulo.EndEmpresa_Id = nf.EndEmpresa_Id " & vbCrLf &
            '    "         AND   Titulo.Cliente_Id = nf.Cliente_Id" & vbCrLf &
            '    "         AND   Titulo.EndCliente_Id = nf.EndCliente_Id " & vbCrLf &
            '    "         AND   Titulo.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
            '    "         AND   Titulo.Serie_Id = nf.Serie_Id " & vbCrLf &
            '    "         AND   Titulo.Nota_Id = nf.Nota_Id " & vbCrLf &
            '    "         AND   ContasAPagar.Provisao = 2" & vbCrLf &
            '    "         And CASE " & vbCrLf &
            '    "               WHEN ContasAPagar.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
            '    "               	 And ContasAPagar.Provisao = 1 THEN 2 " & vbCrLf &
            '    "               ELSE ContasAPagar.Provisao" & vbCrLf &
            '    "             END in(2,3))" & vbCrLf &
            '    "OR (EncargoLiquido.Total =  TitulosBaixados.Total))" & vbCrLf
            '    '"OR (ISNULL(EncargoLiquido.Total,0) =  ISNULL(TitulosBaixados.Total,0)))" & vbCrLf

            'End If

            SqlT &= " AND (" & vbCrLf &
                   "         nf.movimento BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
                   "     ) "

            '"      or Titulos.UsuarioInclusaoData BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf & _

            parametros &= "Movimento Fiscal entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text & vbCrLf
            parametros &= "Data Base em " & txtPeriodoFinalConsultaTitulos.Text

            If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                SqlT &= "and titulos.Pedido = " & txtPedido.Text & "" & vbCrLf
                parametros &= "Pedido: " & txtPedido.Text
            End If

            If RadAtivos.Checked Then
                If chkPrevisao.Visible And chkProvisao.Visible Then
                    If chkPrevisao.Checked And chkProvisao.Checked Then
                        SqlT &= " and (" & Provisao & " = 2 OR " & Provisao & " = 3) " & vbCrLf
                        parametros &= "Titulos Previsionados, Provisionados" & vbCrLf
                    ElseIf chkPrevisao.Checked Then
                        SqlT &= " and ((" & Provisao & " = 2) or (" & Provisao & " = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                        parametros &= "Titulos Previsionados" & vbCrLf
                    ElseIf chkProvisao.Checked Then
                        SqlT &= " and " & Provisao & " = 3 " & vbCrLf
                        parametros &= "Titulos Provisionados" & vbCrLf
                    Else
                        SqlT &= " and ((" & Provisao & " = 2) or (" & Provisao & " = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                        parametros &= "Titulos Previsionados" & vbCrLf
                    End If
                Else
                    SqlT &= " and ((" & Provisao & " = 2) or (" & Provisao & " = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                    parametros &= "Titulos Previsionados" & vbCrLf
                End If
                parametros &= "Titulos Abertos" & vbCrLf
            End If

        Else
            If RadAtivos.Checked Then
                If chkPrevisao.Visible And chkProvisao.Visible Then
                    If chkPrevisao.Checked And chkProvisao.Checked Then
                        SqlT &= " and (Titulos.Provisao = 2 OR Titulos.Provisao = 3 "
                        If (Excel) Then
                            SqlT &= " OR Titulos.Provisao = 1 " & vbCrLf
                        End If
                        SqlT &= " ) " & vbCrLf
                        parametros &= "Titulos Previsionados, Provisionados" & vbCrLf
                    ElseIf chkPrevisao.Checked Then
                        SqlT &= " and ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                        parametros &= "Titulos Previsionados" & vbCrLf
                    ElseIf chkProvisao.Checked Then
                        SqlT &= " and Titulos.Provisao = 3 " & vbCrLf
                        parametros &= "Titulos Provisionados" & vbCrLf
                    Else
                        SqlT &= " and ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                        parametros &= "Titulos Previsionados" & vbCrLf
                    End If
                Else
                    SqlT &= " and ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                    parametros &= "Titulos Previsionados" & vbCrLf
                End If
                parametros &= "Titulos Abertos" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                SqlT &= "and titulos.Pedido = " & txtPedido.Text & "" & vbCrLf
                parametros &= "Pedido: " & txtPedido.Text
            ElseIf IsDate(txtPeriodoInicialConsultaTitulos.Text) AndAlso IsDate(txtPeriodoFinalConsultaTitulos.Text) AndAlso chkPeriodo.Checked Then
                If radMovimento.Checked Then
                    SqlT &= "AND (CASE " & vbCrLf
                    SqlT &= "        WHEN  Titulos.provisao = 1 " & vbCrLf
                    SqlT &= "		    then Titulos.Baixa " & vbCrLf
                    SqlT &= "            else Titulos.Movimento " & vbCrLf
                    SqlT &= "         END BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "')" & vbCrLf
                    parametros &= "Movimento entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text & vbCrLf
                ElseIf radVencimento.Checked Then
                    SqlT &= " AND Titulos.Prorrogacao BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf
                    parametros &= "Vecimentos entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text & vbCrLf
                End If
            End If
        End If

        If RadBaixados.Checked Then
            SqlT &= " and Titulos.Provisao = 1 " & vbCrLf
            parametros &= "Titulos Baixados" & vbCrLf
        End If

        If ddlMoeda.SelectedIndex > 0 Then
            SqlT &= " and Titulos.Moeda = " & ddlMoeda.SelectedValue & vbCrLf
            parametros &= "Moeda: " & ddlMoeda.SelectedItem.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlUnidadeConsultaTitulos.SelectedValue) Then
            SqlT &= " and Titulos.UnidadeDeNegocio = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            If chkConsolidarEmpresa.Checked Then
                SqlT &= " and Titulos.Empresa in(" & EmpresasConsolidas & ")" & vbCrLf
            Else
                SqlT &= " and Titulos.Empresa = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            " and Titulos.EndEmpresa = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            If chkConsolidarCliente.Checked Then
                SqlT &= " and left(Titulos.Cliente,8) = '" & txtCodigoCliente.Value.Split("-")(0).Substring(0, 8) & "'" & vbCrLf
            Else
                SqlT &= " and Titulos.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                            " and Titulos.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            End If
        End If

        If ddlCarteiraDoTitulo.SelectedIndex > 0 Then
            SqlT &= " and Titulos.CarteiraDoTitulo = '" & ddlCarteiraDoTitulo.SelectedValue & "'" & vbCrLf
            parametros &= "Carteira: " & ddlCarteiraDoTitulo.SelectedItem.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlBanco.SelectedValue) Then
            SqlT &= " and Titulos.BancoPagador = " & ddlBanco.SelectedValue & vbCrLf
            parametros &= "Banco: " & ddlBanco.SelectedItem.Text & vbCrLf
        End If

        If ddlFinalidadeFinanceira.SelectedIndex > 0 Then
            SqlT &= " and Titulos.Carteira = '" & ddlFinalidadeFinanceira.SelectedValue & "'" & vbCrLf
            parametros &= "Finalidade Financeira: " & ddlFinalidadeFinanceira.SelectedItem.Text & vbCrLf
        End If

        If radDataBase.Checked Then
            SqlT &= " and isnull(nf.movimento, titulos.UsuarioInclusaoData)  <='" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
        End If

        If radDataBase.Checked Then
            SqlT &= "UNION ALL "

            'Devolucoes integrais - Se for excel as colunas são diferentes
            If Excel Then

                SqlT &= "  SELECT  convert(varchar,Nota.Nota_Id) AS Titulo, "
                SqlT &= "          '' AS Provisao,"
                SqlT &= "          '' AS ContaContabilPagadora,"
                SqlT &= "          '' AS NomeDaConta,"
                SqlT &= "          '' AS CarteiraFinanceira,"
                SqlT &= "          '' AS Descricaocf,"
                SqlT &= "          '' AS CarteiraDoTitulo,"
                SqlT &= "          '' AS Descricaoct,"
                SqlT &= "          '' AS Encargo,"
                SqlT &= "          '' AS NomeEncargo,"
                SqlT &= "          0  AS Moeda,"
                SqlT &= "          CONVERT(VARCHAR,nOTA.Movimento, 103) AS Movimento,"
                SqlT &= "          CONVERT(VARCHAR,nOTA.Movimento, 103) AS Vencimento,"
                SqlT &= "          '' AS Prorrogacao,"
                SqlT &= "          '' AS Baixa,"
                SqlT &= "          Empresa.Reduzido AS ReduzidoEmpresa,"
                SqlT &= "          Empresa.Nome AS NomeDaEmpresa,"
                SqlT &= "          Empresa.Cidade AS CidadeDaEmpresa,"
                SqlT &= "          '' AS GrupodeContas,"
                SqlT &= "          '' AS NomeGrupodeContas,"
                SqlT &= "          Cliente.Reduzido AS Fornecedor,"
                SqlT &= "          Cliente.Nome AS NomeDoFornecedor,"
                SqlT &= "          '' AS ContaContabilEncargo,"
                SqlT &= "          '' AS NomeContaEncargo,"
                SqlT &= "          PedidoNota.Pedido_Id as Pedido,"
                SqlT &= "          SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0)) AS ValorDoDocumento,"
                SqlT &= "          0 as Descontos,"
                SqlT &= "          0 as Deducoes,"
                SqlT &= "          0 as Juros,"
                SqlT &= "          0 as Acrescimos,"
                SqlT &= "         SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0)) as ValorLiquido,"
                SqlT &= "         CONCAT('REF. NF ',Nota.Nota_Id,' - ',Nota.Serie_Id,' - ',Nota.EntradaSaida_Id,' DE ',ProdutoNota.Nome,'- Pedido ',PedidoNota.Pedido_Id) as Historico,"
                SqlT &= "          '' as Observacoes,"
                SqlT &= "          '' as Destinatario,"
                SqlT &= "          '' as NomeDoDestinatario,"
                SqlT &= "          ProdutoNota.Produto_Id as Produto,"
                SqlT &= "          ProdutoNota.Nome as NomeProduto,"
                SqlT &= "   (SELECT SUM(EncargoLiquido.Valor) FROM  NotasFiscaisXItens AS ItensNota     "
                SqlT &= "    INNER Join NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNota.Empresa_Id    "
                SqlT &= "    And EncargoLiquido.EndEmpresa_Id = ItensNota.EndEmpresa_Id     "
                SqlT &= "    And EncargoLiquido.Cliente_Id = ItensNota.Cliente_Id     "
                SqlT &= "    And EncargoLiquido.EndCliente_Id = ItensNota.EndCliente_Id     "
                SqlT &= "    And EncargoLiquido.EntradaSaida_Id = ItensNota.EntradaSaida_Id    "
                SqlT &= "    And EncargoLiquido.Serie_Id = ItensNota.Serie_Id   "
                SqlT &= "    And EncargoLiquido.Nota_Id = ItensNota.Nota_Id    "
                SqlT &= "    And EncargoLiquido.Produto_Id = ItensNota.Produto_Id     "
                SqlT &= "    And EncargoLiquido.Encargo_Id = 'LIQUIDO'     "
                SqlT &= "    WHERE ItensNota.Empresa_Id = Nota.Empresa_Id     "
                SqlT &= "    And ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id     "
                SqlT &= "    And ItensNota.Cliente_Id = Nota.Cliente_Id    "
                SqlT &= "    And ItensNota.EndCliente_Id = Nota.EndCliente_Id     "
                SqlT &= "    And ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id    "
                SqlT &= "    And ItensNota.Serie_Id = Nota.Serie_Id    "
                SqlT &= "    And ItensNota.Nota_Id = Nota.Nota_Id     "
                SqlT &= "   ) AS Total,     "
                SqlT &= "       (SELECT ISNULL(SUM(CASE"
                SqlT &= "                            WHEN TipoDeLancamento = 'E'"
                SqlT &= "                              THEN ISNULL(pXi.TotalMoeda, 0) * - 1"
                SqlT &= "                              ELSE ISNULL(pXi.TotalMoeda, 0)"
                SqlT &= "                          END),0)"
                SqlT &= "          FROM Pedidos"
                SqlT &= "          LEFT JOIN PedidoXItemxLancamento AS pXi"
                SqlT &= "            ON Pedidos.Pedido_Id     = pXi.Pedido_Id"
                SqlT &= "           AND Pedidos.Empresa_Id    = pXi.Empresa_Id"
                SqlT &= "           AND Pedidos.EndEmpresa_Id = pXi.EndEmpresa_Id	"
                SqlT &= "         WHERE Pedidos.Pedido_Id     = PedidoNota.Pedido_Id"
                SqlT &= "           AND Pedidos.Empresa_Id    = PedidoNota.Empresa_Id"
                SqlT &= "           AND Pedidos.EndEmpresa_Id = PedidoNota.EndEmpresa_Id) AS TotalEmDolar,"
                SqlT &= "           Nota.Nota_Id as Nota,"
                SqlT &= "           CONVERT(VARCHAR,Nota.Movimento,103) as MovimentoNota,"
                SqlT &= "           '' as ContratoBanco,"
                SqlT &= "           '' as RegistroMestre,"
                SqlT &= "           SoPedido.Classe,"
                SqlT &= "           '' as ContraPartida,"
                SqlT &= "           '' as DescricaoContrapartida,"
                SqlT &= "           '' as TipoDaConta,"
                SqlT &= "	   isnull((SELECT top 1 raz.Conta_id "
                SqlT &= "	      from razao raz "
                SqlT &= "			where raz.empresa_id    = Nota.Empresa_Id "
                SqlT &= "			and raz.endempresa_id   = Nota.EndEmpresa_Id "
                SqlT &= "			and raz.cliente_nf      = Nota.Cliente_Id "
                SqlT &= "			and raz.endcliente_nf   = Nota.EndCliente_Id "
                SqlT &= "			and raz.entradasaida_nf = Nota.EntradaSaida_Id "
                SqlT &= "			and raz.serie_nf        = Nota.Serie_Id "
                SqlT &= "			and raz.numero_nf       = Nota.Nota_Id "
                SqlT &= "			and raz.Encargo_NF = 'PRODUTO'),'') AS ContrapartidaNota,"
                SqlT &= "	   isnull((SELECT top 1 raz.Conta_id as Titulo "
                SqlT &= "	      from razao raz "
                SqlT &= "			where raz.empresa_id    = Nota.Empresa_Id "
                SqlT &= "			and raz.endempresa_id   = Nota.EndEmpresa_Id "
                SqlT &= "			and raz.cliente_nf      = Nota.Cliente_Id"
                SqlT &= "			and raz.endcliente_nf   = Nota.EndCliente_Id "
                SqlT &= "			and raz.entradasaida_nf = Nota.EntradaSaida_Id "
                SqlT &= "			and raz.serie_nf        = Nota.Serie_Id "
                SqlT &= "			and raz.numero_nf       = Nota.Nota_Id "
                SqlT &= "			and raz.Encargo_NF = 'PRODUTO'),'') AS DescricaoContrapartidaNota,'' Tipo"

            Else

                SqlT &= "  SELECT  Empresa.Reduzido AS ReduzidoEmpresa,"
                SqlT &= "          Empresa.Cliente_Id as Empresa,"
                SqlT &= "          Empresa.Endereco_Id as EndEmpresa,"
                SqlT &= "          Empresa.Nome AS NomeEmpresa,"
                SqlT &= "          Empresa.Cidade AS CidadeEmpresa,"
                SqlT &= "          Empresa.Estado AS EstadoEmpresa,"
                SqlT &= "          convert(varchar,Nota.Nota_Id) AS Titulo,"
                SqlT &= "          convert(nvarchar,Nota.Nota_Id) + ' ' + CASE"
                SqlT &= "														 WHEN isnull(PedidoNota.Moeda, 1) = 1 THEN '(NOTA) R$'"
                SqlT &= "													     ELSE '(NOTA) U$'"
                SqlT &= "                                                 END Registro,"
                SqlT &= "          '' AS Faturamento,"
                SqlT &= "          '' AS Lote,"
                SqlT &= "          0 AS LoteTotal,"
                SqlT &= "          0 AS LoteEntregue,"
                SqlT &= "          PedidoNota.Pedido_Id as Pedido,"
                SqlT &= "          Cliente.Cliente_Id as Cliente,"
                SqlT &= "          Cliente.Nome AS NomeCliente,"
                SqlT &= "          Nota.Movimento,"
                SqlT &= "          Nota.Movimento AS Vencimento,"
                SqlT &= "          Nota.Movimento AS Baixa,"
                SqlT &= "          PedidoNota.Carteira,"
                SqlT &= "          CarteiraPedido.Descricao AS NomeCarteira,"
                SqlT &= "          CONCAT('REF. NF ',Nota.Nota_Id,' - ',Nota.Serie_Id,' - ',Nota.EntradaSaida_Id,' DE ',ProdutoNota.Nome,'- Pedido ',PedidoNota.Pedido_Id) as Historico,"
                SqlT &= "          '' as solicitacao,"
                SqlT &= "          ISNULL(PedidoNota.PedidoEfetivo, 0) PedidoEfetivo,"
                SqlT &= "          '' AS Provisao,"
                SqlT &= "          SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0)) AS ValorLiquido,"
                SqlT &= "          0 AS MoedaValorLiquido,"
                SqlT &= "          '' AS tipo"

            End If

            SqlT &= "   FROM NotasFiscais as Nota "
            SqlT &= ""
            SqlT &= "   INNER JOIN Clientes AS Empresa ON Empresa.Cliente_Id  = Nota.Empresa_Id"
            SqlT &= "   AND Empresa.Endereco_Id = Nota.EndEmpresa_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN SubOperacoes AS SoNota ON SoNota.Operacao_Id  = Nota.Operacao"
            SqlT &= "   AND SoNota.SubOperacoes_Id = Nota.SubOperacao"
            SqlT &= "   AND SoNota.Financeiro = 'S'"
            SqlT &= ""
            SqlT &= "   INNER JOIN Clientes AS Cliente ON Cliente.Cliente_Id  = Nota.Cliente_Id"
            SqlT &= "   AND Cliente.Endereco_Id = Nota.EndCliente_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN Pedidos as PedidoNota on PedidoNota.Empresa_Id = Nota.Empresa_Id"
            SqlT &= "   AND PedidoNota.EndEmpresa_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND PedidoNota.Cliente = Nota.Cliente_Id"
            SqlT &= "   AND PedidoNota.EndCliente = Nota.EndCliente_Id"
            SqlT &= "   AND PedidoNota.Pedido_Id = Nota.Pedido"
            SqlT &= ""
            SqlT &= "   INNER JOIN SubOperacoes AS SoPedido ON SoPedido.Operacao_Id  = PedidoNota.Operacao"
            SqlT &= "   AND SoPedido.SubOperacoes_Id = PedidoNota.SubOperacao"
            SqlT &= ""
            SqlT &= "   INNER JOIN Carteira as CarteiraPedido on CarteiraPedido.Carteira_Id = PedidoNota.Carteira"
            SqlT &= ""
            SqlT &= "   INNER JOIN NotasFiscaisXItens AS ItensNota ON ItensNota.Empresa_Id = Nota.Empresa_Id"
            SqlT &= "   AND ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND ItensNota.Cliente_Id = Nota.Cliente_Id"
            SqlT &= "   AND ItensNota.EndCliente_Id = Nota.EndCliente_Id"
            SqlT &= "   AND ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id"
            SqlT &= "   AND ItensNota.Serie_Id = Nota.Serie_Id"
            SqlT &= "   AND ItensNota.Nota_Id = Nota.Nota_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN Produtos AS ProdutoNota ON ProdutoNota.Produto_Id = ItensNota.Produto_Id"
            SqlT &= "   INNER JOIN NotaFiscalDevolucaoXNotaFiscal AS Devolucoes on Devolucoes.EmpresaDevolucao_Id = Nota.Empresa_Id"
            SqlT &= "   AND Devolucoes.EndEmpresaDevolucao_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND Devolucoes.ClienteDevolucao_Id = Nota.Cliente_Id"
            SqlT &= "   AND Devolucoes.EndClienteDevolucao_Id = Nota.EndCliente_Id"
            SqlT &= "   AND Devolucoes.Produto_Id = ItensNota.Produto_Id"
            SqlT &= "   AND Devolucoes.EntradaSaida_Id = Nota.EntradaSaida_Id"
            SqlT &= "   AND Devolucoes.Serie_Id = Nota.Serie_Id"
            SqlT &= "   AND Devolucoes.Nota_Id = Nota.Nota_Id"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscais AS NotaDevolucao ON NotaDevolucao.Empresa_Id = Devolucoes.EmpresaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.EndEmpresa_Id = Devolucoes.EndEmpresaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Cliente_Id = Devolucoes.ClienteDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.EndCliente_Id = Devolucoes.EndClienteDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.EntradaSaida_Id = Devolucoes.EntradaSaidaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Serie_Id = Devolucoes.SerieDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Nota_Id = Devolucoes.NotaDevolucao_Id"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscaisXItens AS ItensNotaDevolucao ON ItensNotaDevolucao.Empresa_Id = NotaDevolucao.Empresa_Id"
            SqlT &= "   AND ItensNotaDevolucao.EndEmpresa_Id = NotaDevolucao.EndEmpresa_Id"
            SqlT &= "   AND ItensNotaDevolucao.Cliente_Id = NotaDevolucao.Cliente_Id"
            SqlT &= "   AND ItensNotaDevolucao.EndCliente_Id = NotaDevolucao.EndCliente_Id"
            SqlT &= "   AND ItensNotaDevolucao.EntradaSaida_Id = NotaDevolucao.EntradaSaida_Id"
            SqlT &= "   AND ItensNotaDevolucao.Serie_Id = NotaDevolucao.Serie_Id"
            SqlT &= "   AND ItensNotaDevolucao.Nota_Id = NotaDevolucao.Nota_Id"
            SqlT &= "   AND ItensNotaDevolucao.CFOP_Id = Devolucoes.CFOPDevolucao_Id"
            SqlT &= "   AND ItensNotaDevolucao.Produto_Id = Devolucoes.Produto_Id"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNotaDevolucao.Empresa_Id"
            SqlT &= "   AND EncargoLiquido.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id"
            SqlT &= "   AND EncargoLiquido.Cliente_Id = ItensNotaDevolucao.Cliente_Id"
            SqlT &= "   AND EncargoLiquido.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id"
            SqlT &= "   AND EncargoLiquido.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id"
            SqlT &= "   AND EncargoLiquido.Serie_Id = ItensNotaDevolucao.Serie_Id"
            SqlT &= "   AND EncargoLiquido.Nota_Id = ItensNotaDevolucao.Nota_Id"
            SqlT &= "   AND EncargoLiquido.Sequencia_id = ItensNotaDevolucao.Sequencia_id"
            SqlT &= "   AND EncargoLiquido.Produto_Id = ItensNotaDevolucao.Produto_Id"
            SqlT &= "   AND EncargoLiquido.Encargo_Id = 'LIQUIDO'"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscaisXEncargos AS EncargoProduto ON EncargoProduto.Empresa_Id = ItensNotaDevolucao.Empresa_Id"
            SqlT &= "   AND EncargoProduto.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id"
            SqlT &= "   AND EncargoProduto.Cliente_Id = ItensNotaDevolucao.Cliente_Id"
            SqlT &= "   AND EncargoProduto.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id"
            SqlT &= "   AND EncargoProduto.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id"
            SqlT &= "   AND EncargoProduto.Serie_Id = ItensNotaDevolucao.Serie_Id"
            SqlT &= "   AND EncargoProduto.Nota_Id = ItensNotaDevolucao.Nota_Id"
            SqlT &= "   AND EncargoProduto.Sequencia_id = ItensNotaDevolucao.Sequencia_id"
            SqlT &= "   AND EncargoProduto.Produto_Id = ItensNotaDevolucao.Produto_Id"
            SqlT &= "   AND EncargoProduto.Encargo_Id = 'PRODUTO'"

            If ddlSafra.SelectedIndex > 0 Then
                SqlT &= "  Where PedidoNota.Safra = '" & ddlSafra.SelectedValue & "' " & vbCrLf
            Else
                SqlT &= "  Where 1 = 1"
            End If

            If op.Length > 0 Then
                SqlT &= "   AND isnull(PedidoNota.Finalidade,0) " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & op & "')"
            End If

            If ddlRepresentante.SelectedIndex > 0 Then
                SqlT &= "  And exists (Select 1" & vbCrLf &
                        "                from Comissoes C" & vbCrLf &
                        "               where C.Empresa_id          = PedidoNota.Empresa_id" & vbCrLf &
                        "                 And C.EndEmpresa_Id       = PedidoNota.EndEmpresa_Id " & vbCrLf &
                        "                 And C.Pedido_id           = PedidoNota.Pedido_id " & vbCrLf &
                        "                 And C.Representante_id    ='" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                        "                 And C.EndRepresentante_Id = " & ddlRepresentante.SelectedValue.Split("-")(1) & ")" & vbCrLf
            End If

            SqlT &= "    AND Nota.Situacao = 1 "
            SqlT &= "   AND NotaDevolucao.Movimento <= '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"

            'Comentei por não estar aparecendo os Titulos Lançados Tirando do Banco para o Caixa e de Transferência entre Contas - 31/03/2016 - Furlan
            '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf & _
            '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf


            SqlT &= " AND Nota.movimento BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf

            If ddlMoeda.SelectedIndex > 0 Then
                SqlT &= " and PedidoNota.Moeda = " & ddlMoeda.SelectedValue & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    SqlT &= " and Empresa.Cliente_Id in(" & EmpresasConsolidas & ")" & vbCrLf
                Else
                    SqlT &= " and Empresa.Cliente_Id = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            " and Empresa.Endereco_Id = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
                End If
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                If chkConsolidarCliente.Checked Then
                    SqlT &= " and left(Cliente.Cliente_Id,8) = '" & txtCodigoCliente.Value.Split("-")(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    SqlT &= " and Cliente.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                            " and Cliente.Endereco_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
                End If
            End If

            If ddlCarteiraDoTitulo.SelectedIndex > 0 Then
                SqlT &= " and PedidoNota.Carteira = '" & ddlCarteiraDoTitulo.SelectedValue & "'" & vbCrLf
            End If

            If RdReceber.Checked Then

                'SqlT &= " AND NOT EXISTS(Select top 1 * from NotaFiscalXTitulo as Titulo "
                'SqlT &= ""
                'SqlT &= "			     INNER JOIN ContasAReceber ON ContasAReceber.Registro_Id = Titulo.Titulo_Id"
                'SqlT &= ""
                'SqlT &= "                WHERE Titulo.Empresa_Id = Nota.Empresa_Id"
                'SqlT &= "			     AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id"
                'SqlT &= "			     AND   Titulo.Cliente_Id = Nota.Cliente_Id"
                'SqlT &= "			     AND   Titulo.EndCliente_Id = Nota.EndCliente_Id"
                'SqlT &= "			     AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id"
                'SqlT &= "		         AND   Titulo.Serie_Id = Nota.Serie_Id"
                'SqlT &= "		         AND   Titulo.Nota_Id = Nota.Nota_Id "
                'SqlT &= "                )"

                SqlT &= "AND Nota.EntradaSaida_Id = 'S' "

            ElseIf RdPagar.Checked Then

                'SqlT &= " AND NOT EXISTS(Select top 1 * from NotaFiscalXTitulo as Titulo "
                'SqlT &= ""
                'SqlT &= "                INNER JOIN ContasAPagar ON ContasAPagar.Registro_Id = Titulo.Titulo_Id"
                'SqlT &= ""
                'SqlT &= "                WHERE Titulo.Empresa_Id = Nota.Empresa_Id"
                'SqlT &= "			     AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id"
                'SqlT &= "			     AND   Titulo.Cliente_Id = Nota.Cliente_Id"
                'SqlT &= "			     AND   Titulo.EndCliente_Id = Nota.EndCliente_Id"
                'SqlT &= "			     AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id"
                'SqlT &= "		         AND   Titulo.Serie_Id = Nota.Serie_Id"
                'SqlT &= "		         AND   Titulo.Nota_Id = Nota.Nota_Id "
                'SqlT &= "                )"

                SqlT &= "AND Nota.EntradaSaida_Id = 'E' "

            End If

            SqlT &= "    AND EXISTS (
				                        SELECT 1 
				                        FROM NotaFiscalXTitulo AS Titulos "

            If RdReceber.Checked Then
                SqlT &= "               LEFT JOIN ContasAReceber as Titulo " & vbCrLf &
                        "                       ON Titulo.Registro_Id       = Titulos.Titulo_Id " & vbCrLf

                If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                    SqlT &= "                       AND Titulo.Situacao in(1,101,102,105) " & vbCrLf
                Else
                    SqlT &= "                       AND Titulo.Situacao in(1,101,102) " & vbCrLf
                End If
            Else
                SqlT &= "               LEFT JOIN ContasAPagar as Titulo " & vbCrLf &
                        "                       On Titulo.Registro_Id       = Titulos.Titulo_Id " & vbCrLf

                If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                    SqlT &= "                       AND Titulo.Situacao in(1,101,102,105) " & vbCrLf
                Else
                    SqlT &= "                       AND Titulo.Situacao in(1,101,102) " & vbCrLf
                End If
            End If

            'Se o titulo ja esta no periodo isola... senao ele adiciona de novo duplicando o valor
            SqlT &= "                   WHERE Titulos.Empresa_Id		= Nota.Empresa_Id
					                        AND Titulos.EndEmpresa_Id	= Nota.EndEmpresa_Id
					                        AND Titulos.Cliente_Id		= Nota.Cliente_Id
					                        AND Titulos.EndCliente_Id	= Nota.EndCliente_Id
					                        AND Titulos.EntradaSaida_Id = Nota.EntradaSaida_Id
					                        AND Titulos.Serie_Id		= Nota.Serie_Id
					                        AND Titulos.Nota_Id			= Nota.Nota_Id
                                            AND ((NOT ( CASE " & vbCrLf &
                "                                           WHEN Titulo.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
                "               	                            AND Titulo.Provisao = 1 THEN 2 " & vbCrLf &
                "                                           ELSE Titulo.Provisao" & vbCrLf &
                "                                       END = 2" & vbCrLf &
                "                                       OR 
                                                        CASE" & vbCrLf &
                "                                           WHEN Titulo.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
                "                                           AND Titulo.Provisao = 1 THEN 2" & vbCrLf &
                "                                       ELSE Titulo.Provisao" & vbCrLf &
                "                                       END = 3)) OR Titulo.Registro_Id IS NULL)
                                    ) " & vbCrLf

            If Excel Then

                SqlT &= ""
                SqlT &= " GROUP BY   Nota.Nota_id,"
                SqlT &= "            Nota.Movimento,"
                SqlT &= "            Empresa.Reduzido,"
                SqlT &= "            Empresa.Nome,"
                SqlT &= "            Empresa.Cidade,"
                SqlT &= "            Cliente.Reduzido,"
                SqlT &= "            Cliente.Nome,"
                SqlT &= "            PedidoNota.Pedido_Id,"
                SqlT &= "            Nota.Serie_Id,"
                SqlT &= "            Nota.EntradaSaida_Id,"
                SqlT &= "            ProdutoNota.Produto_Id,"
                SqlT &= "            ProdutoNota.Nome,"
                SqlT &= "            SoPedido.Classe,"
                SqlT &= "            PedidoNota.Empresa_Id,"
                SqlT &= "            PedidoNota.EndEmpresa_Id,"
                SqlT &= "            Nota.Empresa_Id,"
                SqlT &= "            Nota.EndEmpresa_Id,"
                SqlT &= "            Nota.Cliente_Id,"
                SqlT &= "            Nota.EndCliente_Id"
            Else

                SqlT &= ""
                SqlT &= " GROUP BY   Nota.Nota_id,"
                SqlT &= "		     Empresa.Reduzido,"
                SqlT &= "		     Empresa.Cliente_Id,"
                SqlT &= "		     Empresa.Endereco_Id,"
                SqlT &= "		     Empresa.Nome,"
                SqlT &= "		     Empresa.Cidade,"
                SqlT &= "		     Empresa.Estado,"
                SqlT &= "		     PedidoNota.Moeda, "
                SqlT &= "	         PedidoNota.Pedido_Id,"
                SqlT &= "		     Cliente.Cliente_Id, "
                SqlT &= "	         Cliente.Nome,"
                SqlT &= "            Nota.Movimento,"
                SqlT &= "            PedidoNota.Carteira,"
                SqlT &= "            CarteiraPedido.Descricao,"
                SqlT &= "            Nota.Serie_Id,"
                SqlT &= "            Nota.EntradaSaida_Id,"
                SqlT &= "            ProdutoNota.Nome,"
                SqlT &= "            PedidoNota.PedidoEfetivo"

            End If



            SqlT &= ""
            SqlT &= " HAVING(SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0))) > 0 "


            'Rodrigo 25/07/2023 - Adicionado esse union para tratar devolucoes na relação, o de cima trata os dentro do periodo
            'Essa parte do script não irá ser consultado, está gerando alguns problemas na VERDE
            'Será analisado - 28/07/2025 GILBERTO
            SqlT &= "UNION "

            'Devolucoes integrais - Se for excel as colunas são diferentes
            If Excel Then

                SqlT &= "  SELECT  convert(varchar,Nota.Nota_Id) AS Titulo, "
                SqlT &= "          '' AS Provisao,"
                SqlT &= "          '' AS ContaContabilPagadora,"
                SqlT &= "          '' AS NomeDaConta,"
                SqlT &= "          '' AS CarteiraFinanceira,"
                SqlT &= "          '' AS Descricaocf,"
                SqlT &= "          '' AS CarteiraDoTitulo,"
                SqlT &= "          '' AS Descricaoct,"
                SqlT &= "          '' AS Encargo,"
                SqlT &= "          '' AS NomeEncargo,"
                SqlT &= "          0  AS Moeda,"
                SqlT &= "          CONVERT(VARCHAR,nOTA.Movimento, 103) AS Movimento,"
                SqlT &= "          CONVERT(VARCHAR,nOTA.Movimento, 103) AS Vencimento,"
                SqlT &= "          '' AS Prorrogacao,"
                SqlT &= "          '' AS Baixa,"
                SqlT &= "          Empresa.Reduzido AS ReduzidoEmpresa,"
                SqlT &= "          Empresa.Nome AS NomeDaEmpresa,"
                SqlT &= "          Empresa.Cidade AS CidadeDaEmpresa,"
                SqlT &= "          '' AS GrupodeContas,"
                SqlT &= "          '' AS NomeGrupodeContas,"
                SqlT &= "          Cliente.Reduzido AS Fornecedor,"
                SqlT &= "          Cliente.Nome AS NomeDoFornecedor,"
                SqlT &= "          '' AS ContaContabilEncargo,"
                SqlT &= "          '' AS NomeContaEncargo,"
                SqlT &= "          PedidoNota.Pedido_Id as Pedido,"
                SqlT &= "          SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0)) AS ValorDoDocumento,"
                SqlT &= "          0 as Descontos,"
                SqlT &= "          0 as Deducoes,"
                SqlT &= "          0 as Juros,"
                SqlT &= "          0 as Acrescimos,"
                SqlT &= "         SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0)) as ValorLiquido,"
                SqlT &= "         CONCAT('REF. NF ',Nota.Nota_Id,' - ',Nota.Serie_Id,' - ',Nota.EntradaSaida_Id,' DE ',ProdutoNota.Nome,'- Pedido ',PedidoNota.Pedido_Id) as Historico,"
                SqlT &= "          '' as Observacoes,"
                SqlT &= "          '' as Destinatario,"
                SqlT &= "          '' as NomeDoDestinatario,"
                SqlT &= "          ProdutoNota.Produto_Id as Produto,"
                SqlT &= "          ProdutoNota.Nome as NomeProduto,"
                SqlT &= "   (SELECT SUM(EncargoLiquido.Valor) FROM  NotasFiscaisXItens AS ItensNota     "
                SqlT &= "    INNER Join NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNota.Empresa_Id    "
                SqlT &= "    And EncargoLiquido.EndEmpresa_Id = ItensNota.EndEmpresa_Id     "
                SqlT &= "    And EncargoLiquido.Cliente_Id = ItensNota.Cliente_Id     "
                SqlT &= "    And EncargoLiquido.EndCliente_Id = ItensNota.EndCliente_Id     "
                SqlT &= "    And EncargoLiquido.EntradaSaida_Id = ItensNota.EntradaSaida_Id    "
                SqlT &= "    And EncargoLiquido.Serie_Id = ItensNota.Serie_Id   "
                SqlT &= "    And EncargoLiquido.Nota_Id = ItensNota.Nota_Id    "
                SqlT &= "    And EncargoLiquido.Produto_Id = ItensNota.Produto_Id     "
                SqlT &= "    And EncargoLiquido.Encargo_Id = 'LIQUIDO'     "
                SqlT &= "    WHERE ItensNota.Empresa_Id = Nota.Empresa_Id     "
                SqlT &= "    And ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id     "
                SqlT &= "    And ItensNota.Cliente_Id = Nota.Cliente_Id    "
                SqlT &= "    And ItensNota.EndCliente_Id = Nota.EndCliente_Id     "
                SqlT &= "    And ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id    "
                SqlT &= "    And ItensNota.Serie_Id = Nota.Serie_Id    "
                SqlT &= "    And ItensNota.Nota_Id = Nota.Nota_Id     "
                SqlT &= "   ) AS Total,     "
                SqlT &= "       (SELECT ISNULL(SUM(CASE"
                SqlT &= "                            WHEN TipoDeLancamento = 'E'"
                SqlT &= "                              THEN ISNULL(pXi.TotalMoeda, 0) * - 1"
                SqlT &= "                              ELSE ISNULL(pXi.TotalMoeda, 0)"
                SqlT &= "                          END),0)"
                SqlT &= "          FROM Pedidos"
                SqlT &= "          LEFT JOIN PedidoXItemxLancamento AS pXi"
                SqlT &= "            ON Pedidos.Pedido_Id     = pXi.Pedido_Id"
                SqlT &= "           AND Pedidos.Empresa_Id    = pXi.Empresa_Id"
                SqlT &= "           AND Pedidos.EndEmpresa_Id = pXi.EndEmpresa_Id	"
                SqlT &= "         WHERE Pedidos.Pedido_Id     = PedidoNota.Pedido_Id"
                SqlT &= "           AND Pedidos.Empresa_Id    = PedidoNota.Empresa_Id"
                SqlT &= "           AND Pedidos.EndEmpresa_Id = PedidoNota.EndEmpresa_Id) AS TotalEmDolar,"
                SqlT &= "           Nota.Nota_Id as Nota,"
                SqlT &= "           CONVERT(VARCHAR,Nota.Movimento,103) as MovimentoNota,"
                SqlT &= "           '' as ContratoBanco,"
                SqlT &= "           '' as RegistroMestre,"
                SqlT &= "           SoPedido.Classe,"
                SqlT &= "           '' as ContraPartida,"
                SqlT &= "           '' as DescricaoContrapartida,"
                SqlT &= "           '' as TipoDaConta,"
                SqlT &= "	   isnull((SELECT top 1 raz.Conta_id "
                SqlT &= "	      from razao raz "
                SqlT &= "			where raz.empresa_id    = Nota.Empresa_Id "
                SqlT &= "			and raz.endempresa_id   = Nota.EndEmpresa_Id "
                SqlT &= "			and raz.cliente_nf      = Nota.Cliente_Id "
                SqlT &= "			and raz.endcliente_nf   = Nota.EndCliente_Id "
                SqlT &= "			and raz.entradasaida_nf = Nota.EntradaSaida_Id "
                SqlT &= "			and raz.serie_nf        = Nota.Serie_Id "
                SqlT &= "			and raz.numero_nf       = Nota.Nota_Id "
                SqlT &= "			and raz.Encargo_NF = 'PRODUTO'),'') AS ContrapartidaNota,"
                SqlT &= "	   isnull((SELECT top 1  raz.Conta_id as Titulo "
                SqlT &= "	      from razao raz "
                SqlT &= "			where raz.empresa_id    = Nota.Empresa_Id "
                SqlT &= "			and raz.endempresa_id   = Nota.EndEmpresa_Id "
                SqlT &= "			and raz.cliente_nf      = Nota.Cliente_Id"
                SqlT &= "			and raz.endcliente_nf   = Nota.EndCliente_Id "
                SqlT &= "			and raz.entradasaida_nf = Nota.EntradaSaida_Id "
                SqlT &= "			and raz.serie_nf        = Nota.Serie_Id "
                SqlT &= "			and raz.numero_nf       = Nota.Nota_Id "
                SqlT &= "			and raz.Encargo_NF = 'PRODUTO'),'') AS DescricaoContrapartidaNota,'' Tipo"

            Else

                SqlT &= "  SELECT  Empresa.Reduzido AS ReduzidoEmpresa,"
                SqlT &= "          Empresa.Cliente_Id as Empresa,"
                SqlT &= "          Empresa.Endereco_Id as EndEmpresa,"
                SqlT &= "          Empresa.Nome AS NomeEmpresa,"
                SqlT &= "          Empresa.Cidade AS CidadeEmpresa,"
                SqlT &= "          Empresa.Estado AS EstadoEmpresa,"
                SqlT &= "          convert(varchar,Nota.Nota_Id) AS Titulo,"
                SqlT &= "          convert(nvarchar,Nota.Nota_Id) + ' ' + CASE"
                SqlT &= "														 WHEN isnull(PedidoNota.Moeda, 1) = 1 THEN '(NOTA) R$'"
                SqlT &= "													     ELSE '(NOTA) U$'"
                SqlT &= "                                                 END Registro,"
                SqlT &= "          '' AS Faturamento,"
                SqlT &= "          '' AS Lote,"
                SqlT &= "          0 AS LoteTotal,"
                SqlT &= "          0 AS LoteEntregue,"
                SqlT &= "          PedidoNota.Pedido_Id as Pedido,"
                SqlT &= "          Cliente.Cliente_Id as Cliente,"
                SqlT &= "          Cliente.Nome AS NomeCliente,"
                SqlT &= "          Nota.Movimento,"
                SqlT &= "          Nota.Movimento AS Vencimento,"
                SqlT &= "          Nota.Movimento AS Baixa,"
                SqlT &= "          PedidoNota.Carteira,"
                SqlT &= "          CarteiraPedido.Descricao AS NomeCarteira,"
                SqlT &= "          CONCAT('REF. NF ',Nota.Nota_Id,' - ',Nota.Serie_Id,' - ',Nota.EntradaSaida_Id,' DE ',ProdutoNota.Nome,'- Pedido ',PedidoNota.Pedido_Id) as Historico,"
                SqlT &= "          '' as solicitacao,"
                SqlT &= "          ISNULL(PedidoNota.PedidoEfetivo, 0) PedidoEfetivo,"
                SqlT &= "          '' AS Provisao,"
                SqlT &= "          SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0)) AS ValorLiquido,"
                SqlT &= "          0 AS MoedaValorLiquido,"
                SqlT &= "          '' AS tipo"

            End If

            SqlT &= "   FROM NotasFiscais as Nota "
            SqlT &= ""
            SqlT &= "   INNER JOIN Clientes AS Empresa ON Empresa.Cliente_Id  = Nota.Empresa_Id"
            SqlT &= "   AND Empresa.Endereco_Id = Nota.EndEmpresa_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN SubOperacoes AS SoNota ON SoNota.Operacao_Id  = Nota.Operacao"
            SqlT &= "   AND SoNota.SubOperacoes_Id = Nota.SubOperacao"
            SqlT &= "   AND SoNota.Financeiro = 'S'"
            SqlT &= ""
            SqlT &= "   INNER JOIN Clientes AS Cliente ON Cliente.Cliente_Id  = Nota.Cliente_Id"
            SqlT &= "   AND Cliente.Endereco_Id = Nota.EndCliente_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN Pedidos as PedidoNota on PedidoNota.Empresa_Id = Nota.Empresa_Id"
            SqlT &= "   AND PedidoNota.EndEmpresa_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND PedidoNota.Cliente = Nota.Cliente_Id"
            SqlT &= "   AND PedidoNota.EndCliente = Nota.EndCliente_Id"
            SqlT &= "   AND PedidoNota.Pedido_Id = Nota.Pedido"
            SqlT &= ""
            SqlT &= "   INNER JOIN SubOperacoes AS SoPedido ON SoPedido.Operacao_Id  = PedidoNota.Operacao"
            SqlT &= "   AND SoPedido.SubOperacoes_Id = PedidoNota.SubOperacao"
            SqlT &= ""
            SqlT &= "   INNER JOIN Carteira as CarteiraPedido on CarteiraPedido.Carteira_Id = PedidoNota.Carteira"
            SqlT &= ""
            SqlT &= "   INNER JOIN NotasFiscaisXItens AS ItensNota ON ItensNota.Empresa_Id = Nota.Empresa_Id"
            SqlT &= "   AND ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND ItensNota.Cliente_Id = Nota.Cliente_Id"
            SqlT &= "   AND ItensNota.EndCliente_Id = Nota.EndCliente_Id"
            SqlT &= "   AND ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id"
            SqlT &= "   AND ItensNota.Serie_Id = Nota.Serie_Id"
            SqlT &= "   AND ItensNota.Nota_Id = Nota.Nota_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN Produtos AS ProdutoNota ON ProdutoNota.Produto_Id = ItensNota.Produto_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN NotaFiscalDevolucaoXNotaFiscal AS Devolucoes on Devolucoes.EmpresaDevolucao_Id = Nota.Empresa_Id"
            SqlT &= "   AND Devolucoes.EndEmpresaDevolucao_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND Devolucoes.ClienteDevolucao_Id = Nota.Cliente_Id"
            SqlT &= "   AND Devolucoes.EndClienteDevolucao_Id = Nota.EndCliente_Id"
            SqlT &= "   AND Devolucoes.Produto_Id = ItensNota.Produto_Id"
            SqlT &= "   AND Devolucoes.EntradaSaida_Id = Nota.EntradaSaida_Id"
            SqlT &= "   AND Devolucoes.Serie_Id = Nota.Serie_Id"
            SqlT &= "   AND Devolucoes.Nota_Id = Nota.Nota_Id"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscais AS NotaDevolucao ON NotaDevolucao.Empresa_Id = Devolucoes.EmpresaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.EndEmpresa_Id = Devolucoes.EndEmpresaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Cliente_Id = Devolucoes.ClienteDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.EndCliente_Id = Devolucoes.EndClienteDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.EntradaSaida_Id = Devolucoes.EntradaSaidaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Serie_Id = Devolucoes.SerieDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Nota_Id = Devolucoes.NotaDevolucao_Id"
            SqlT &= "   AND NotaDevolucao.Movimento <= '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscaisXItens AS ItensNotaDevolucao ON ItensNotaDevolucao.Empresa_Id = NotaDevolucao.Empresa_Id"
            SqlT &= "   AND ItensNotaDevolucao.EndEmpresa_Id = NotaDevolucao.EndEmpresa_Id"
            SqlT &= "   AND ItensNotaDevolucao.Cliente_Id = NotaDevolucao.Cliente_Id"
            SqlT &= "   AND ItensNotaDevolucao.EndCliente_Id = NotaDevolucao.EndCliente_Id"
            SqlT &= "   AND ItensNotaDevolucao.EntradaSaida_Id = NotaDevolucao.EntradaSaida_Id"
            SqlT &= "   AND ItensNotaDevolucao.Serie_Id = NotaDevolucao.Serie_Id"
            SqlT &= "   AND ItensNotaDevolucao.Nota_Id = NotaDevolucao.Nota_Id"
            SqlT &= "   AND ItensNotaDevolucao.CFOP_Id = Devolucoes.CFOPDevolucao_Id"
            SqlT &= "   AND ItensNotaDevolucao.Produto_Id = Devolucoes.Produto_Id"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNotaDevolucao.Empresa_Id"
            SqlT &= "   AND EncargoLiquido.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id"
            SqlT &= "   AND EncargoLiquido.Cliente_Id = ItensNotaDevolucao.Cliente_Id"
            SqlT &= "   AND EncargoLiquido.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id"
            SqlT &= "   AND EncargoLiquido.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id"
            SqlT &= "   AND EncargoLiquido.Serie_Id = ItensNotaDevolucao.Serie_Id"
            SqlT &= "   AND EncargoLiquido.Nota_Id = ItensNotaDevolucao.Nota_Id"
            SqlT &= "   AND EncargoLiquido.Sequencia_id = ItensNotaDevolucao.Sequencia_id"
            SqlT &= "   AND EncargoLiquido.Produto_Id = ItensNotaDevolucao.Produto_Id"
            SqlT &= "   AND EncargoLiquido.Encargo_Id = 'LIQUIDO'"
            SqlT &= ""
            SqlT &= "   LEFT JOIN NotasFiscaisXEncargos AS EncargoProduto ON EncargoProduto.Empresa_Id = ItensNotaDevolucao.Empresa_Id"
            SqlT &= "   AND EncargoProduto.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id"
            SqlT &= "   AND EncargoProduto.Cliente_Id = ItensNotaDevolucao.Cliente_Id"
            SqlT &= "   AND EncargoProduto.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id"
            SqlT &= "   AND EncargoProduto.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id"
            SqlT &= "   AND EncargoProduto.Serie_Id = ItensNotaDevolucao.Serie_Id"
            SqlT &= "   AND EncargoProduto.Nota_Id = ItensNotaDevolucao.Nota_Id"
            SqlT &= "   AND EncargoProduto.Sequencia_id = ItensNotaDevolucao.Sequencia_id"
            SqlT &= "   AND EncargoProduto.Produto_Id = ItensNotaDevolucao.Produto_Id"
            SqlT &= "   AND EncargoProduto.Encargo_Id = 'PRODUTO'"

            If ddlSafra.SelectedIndex > 0 Then
                SqlT &= "  Where PedidoNota.Safra = '" & ddlSafra.SelectedValue & "' " & vbCrLf
            Else
                SqlT &= "  Where 1 = 1"
            End If

            SqlT &= "  /*  ESSA CONSULTA NÃO SERÁ REALIZA - PODE ESTAR GERANDO VALORES DUPLICADOS, PQ O SCRIPT ABAIXO JÁ TRATA DEVOLUÇÃO */ "
            SqlT &= "  AND 1=2 "

            If op.Length > 0 Then
                SqlT &= "   AND isnull(PedidoNota.Finalidade,0) " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & op & "')"
            End If

            If ddlRepresentante.SelectedIndex > 0 Then
                SqlT &= "  And exists (Select 1" & vbCrLf &
                        "                from Comissoes C" & vbCrLf &
                        "               where C.Empresa_id          = PedidoNota.Empresa_id" & vbCrLf &
                        "                 And C.EndEmpresa_Id       = PedidoNota.EndEmpresa_Id " & vbCrLf &
                        "                 And C.Pedido_id           = PedidoNota.Pedido_id " & vbCrLf &
                        "                 And C.Representante_id    ='" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                        "                 And C.EndRepresentante_Id = " & ddlRepresentante.SelectedValue.Split("-")(1) & ")" & vbCrLf
            End If

            SqlT &= "    AND Nota.Situacao = 1 "

            'Comentei por não estar aparecendo os Titulos Lançados Tirando do Banco para o Caixa e de Transferência entre Contas - 31/03/2016 - Furlan
            '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf & _
            '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf


            SqlT &= " AND Nota.movimento BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf

            If ddlMoeda.SelectedIndex > 0 Then
                SqlT &= " and PedidoNota.Moeda = " & ddlMoeda.SelectedValue & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    SqlT &= " and Empresa.Cliente_Id in(" & EmpresasConsolidas & ")" & vbCrLf
                Else
                    SqlT &= " and Empresa.Cliente_Id = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            " and Empresa.Endereco_Id = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
                End If
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                If chkConsolidarCliente.Checked Then
                    SqlT &= " and left(Cliente.Cliente_Id,8) = '" & txtCodigoCliente.Value.Split("-")(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    SqlT &= " and Cliente.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                            " and Cliente.Endereco_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
                End If
            End If

            If ddlCarteiraDoTitulo.SelectedIndex > 0 Then
                SqlT &= " and PedidoNota.Carteira = '" & ddlCarteiraDoTitulo.SelectedValue & "'" & vbCrLf
            End If

            If RdReceber.Checked Then

                'SqlT &= " AND NOT EXISTS(Select top 1 * from NotaFiscalXTitulo as Titulo "
                'SqlT &= ""
                'SqlT &= "			     INNER JOIN ContasAReceber ON ContasAReceber.Registro_Id = Titulo.Titulo_Id"
                'SqlT &= ""
                'SqlT &= "                WHERE Titulo.Empresa_Id = Nota.Empresa_Id"
                'SqlT &= "			     AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id"
                'SqlT &= "			     AND   Titulo.Cliente_Id = Nota.Cliente_Id"
                'SqlT &= "			     AND   Titulo.EndCliente_Id = Nota.EndCliente_Id"
                'SqlT &= "			     AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id"
                'SqlT &= "		         AND   Titulo.Serie_Id = Nota.Serie_Id"
                'SqlT &= "		         AND   Titulo.Nota_Id = Nota.Nota_Id "
                'SqlT &= "                )"

                SqlT &= "AND Nota.EntradaSaida_Id = 'S' "

            ElseIf RdPagar.Checked Then

                'SqlT &= " AND NOT EXISTS(Select top 1 * from NotaFiscalXTitulo as Titulo "
                'SqlT &= ""
                'SqlT &= "                INNER JOIN ContasAPagar ON ContasAPagar.Registro_Id = Titulo.Titulo_Id"
                'SqlT &= ""
                'SqlT &= "                WHERE Titulo.Empresa_Id = Nota.Empresa_Id"
                'SqlT &= "			     AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id"
                'SqlT &= "			     AND   Titulo.Cliente_Id = Nota.Cliente_Id"
                'SqlT &= "			     AND   Titulo.EndCliente_Id = Nota.EndCliente_Id"
                'SqlT &= "			     AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id"
                'SqlT &= "		         AND   Titulo.Serie_Id = Nota.Serie_Id"
                'SqlT &= "		         AND   Titulo.Nota_Id = Nota.Nota_Id "
                'SqlT &= "                )"

                SqlT &= "AND Nota.EntradaSaida_Id = 'E' "

            End If

            SqlT &= "    AND EXISTS (
				                        SELECT 1 
				                        FROM NotaFiscalXTitulo AS Titulos "

            If RdReceber.Checked Then
                SqlT &= "               LEFT JOIN ContasAReceber as Titulo " & vbCrLf &
                        "                      ON Titulo.Registro_Id       = Titulos.Titulo_Id " & vbCrLf

                If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                    SqlT &= "                      AND Titulo.Situacao in(1,101,102,105) " & vbCrLf
                Else
                    SqlT &= "                      AND Titulo.Situacao in(1,101,102) " & vbCrLf
                End If

            Else
                SqlT &= "               LEFT JOIN ContasAPagar as Titulo " & vbCrLf &
                        "                      ON Titulo.Registro_Id       = Titulos.Titulo_Id " & vbCrLf

                If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                    SqlT &= "                      AND Titulo.Situacao in(1,101,102,105) " & vbCrLf
                Else
                    SqlT &= "                      AND Titulo.Situacao in(1,101,102) " & vbCrLf
                End If

            End If

            'Se o titulo ja esta no periodo isola... senao ele adiciona de novo duplicando o valor
            SqlT &= "                   WHERE Titulos.Empresa_Id		= Nota.Empresa_Id
					                        AND Titulos.EndEmpresa_Id	= Nota.EndEmpresa_Id
					                        AND Titulos.Cliente_Id		= Nota.Cliente_Id
					                        AND Titulos.EndCliente_Id	= Nota.EndCliente_Id
					                        AND Titulos.EntradaSaida_Id = Nota.EntradaSaida_Id
					                        AND Titulos.Serie_Id		= Nota.Serie_Id
					                        AND Titulos.Nota_Id			= Nota.Nota_Id
                                            AND ((NOT ( CASE " & vbCrLf &
                "                                           WHEN Titulo.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
                "               	                            AND Titulo.Provisao = 1 THEN 2 " & vbCrLf &
                "                                           ELSE Titulo.Provisao" & vbCrLf &
                "                                       END = 2" & vbCrLf &
                "                                       OR 
                                                        CASE" & vbCrLf &
                "                                           WHEN Titulo.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf &
                "                                           AND Titulo.Provisao = 1 THEN 2" & vbCrLf &
                "                                       ELSE Titulo.Provisao" & vbCrLf &
                "                                       END = 3)) OR Titulo.Registro_Id IS NULL)
                                    ) " & vbCrLf

            If Excel Then

                SqlT &= ""
                SqlT &= " GROUP BY   Nota.Nota_id,"
                SqlT &= "            Nota.Movimento,"
                SqlT &= "            Empresa.Reduzido,"
                SqlT &= "            Empresa.Nome,"
                SqlT &= "            Empresa.Cidade,"
                SqlT &= "            Cliente.Reduzido,"
                SqlT &= "            Cliente.Nome,"
                SqlT &= "            PedidoNota.Pedido_Id,"
                SqlT &= "            Nota.Serie_Id,"
                SqlT &= "            Nota.EntradaSaida_Id,"
                SqlT &= "            ProdutoNota.Produto_Id,"
                SqlT &= "            ProdutoNota.Nome,"
                SqlT &= "            SoPedido.Classe,"
                SqlT &= "            PedidoNota.Empresa_Id,"
                SqlT &= "            PedidoNota.EndEmpresa_Id,"
                SqlT &= "            Nota.Empresa_Id,"
                SqlT &= "            Nota.EndEmpresa_Id,"
                SqlT &= "            Nota.Cliente_Id,"
                SqlT &= "            Nota.EndCliente_Id"
            Else

                SqlT &= ""
                SqlT &= " GROUP BY   Nota.Nota_id,"
                SqlT &= "		     Empresa.Reduzido,"
                SqlT &= "		     Empresa.Cliente_Id,"
                SqlT &= "		     Empresa.Endereco_Id,"
                SqlT &= "		     Empresa.Nome,"
                SqlT &= "		     Empresa.Cidade,"
                SqlT &= "		     Empresa.Estado,"
                SqlT &= "		     PedidoNota.Moeda, "
                SqlT &= "	         PedidoNota.Pedido_Id,"
                SqlT &= "		     Cliente.Cliente_Id, "
                SqlT &= "	         Cliente.Nome,"
                SqlT &= "            Nota.Movimento,"
                SqlT &= "            PedidoNota.Carteira,"
                SqlT &= "            CarteiraPedido.Descricao,"
                SqlT &= "            Nota.Serie_Id,"
                SqlT &= "            Nota.EntradaSaida_Id,"
                SqlT &= "            ProdutoNota.Nome,"
                SqlT &= "            PedidoNota.PedidoEfetivo"

            End If



            SqlT &= ""
            SqlT &= " HAVING(SUM(IIF(ISNULL(Devolucoes.ValorLiquido,0) = 0, Devolucoes.Valor, Devolucoes.ValorLiquido)) - SUM(ISNULL(EncargoLiquido.Valor,0))) > 0 "


            'Rodrigo 03/05/2023 - Adiciona esse union para tratar notas com saldo aberto que nao tem titulo da diferenca
            SqlT &= "UNION ALL "

            If Excel Then

                SqlT &= "  SELECT  convert(varchar,Nota.Nota_Id) AS Titulo, "
                SqlT &= "          '' AS Provisao,"
                SqlT &= "          '' AS ContaContabilPagadora,"
                SqlT &= "          '' AS NomeDaConta,"
                SqlT &= "          '' AS CarteiraFinanceira,"
                SqlT &= "          '' AS Descricaocf,"
                SqlT &= "          '' AS CarteiraDoTitulo,"
                SqlT &= "          '' AS Descricaoct,"
                SqlT &= "          '' AS Encargo,"
                SqlT &= "          '' AS NomeEncargo,"
                SqlT &= "          0  AS Moeda,"
                SqlT &= "          CONVERT(VARCHAR,nOTA.Movimento, 103) AS Movimento,"
                SqlT &= "          CONVERT(VARCHAR,nOTA.Movimento, 103) AS Vencimento,"
                SqlT &= "          '' AS Prorrogacao,"
                SqlT &= "          '' AS Baixa,"
                SqlT &= "          Empresa.Reduzido AS ReduzidoEmpresa,"
                SqlT &= "          Empresa.Nome AS NomeDaEmpresa,"
                SqlT &= "          Empresa.Cidade AS CidadeDaEmpresa,"
                SqlT &= "          '' AS GrupodeContas,"
                SqlT &= "          '' AS NomeGrupodeContas,"
                SqlT &= "          Cliente.Reduzido AS Fornecedor,"
                SqlT &= "          Cliente.Nome AS NomeDoFornecedor,"
                SqlT &= "          '' AS ContaContabilEncargo,"
                SqlT &= "          '' AS NomeContaEncargo,"
                SqlT &= "          PedidoNota.Pedido_Id as Pedido,"
                SqlT &= "          ((SUM(ISNULL(EncargoLiquidoNotaProduto.Valor, 0)) - ISNULL(Devolucoes.ValorPeriodo, 0)) - (ISNULL(TitulosBaixados.Total, 0) + ISNULL(Devolucoes.ValorFuturo, 0))) AS ValorDoDocumento,"
                SqlT &= "          0 as Descontos,"
                SqlT &= "          0 as Deducoes,"
                SqlT &= "          0 as Juros,"
                SqlT &= "          0 as Acrescimos,"
                SqlT &= "         ((SUM(ISNULL(EncargoLiquidoNotaProduto.Valor, 0)) - ISNULL(Devolucoes.ValorPeriodo, 0)) - (ISNULL(TitulosBaixados.Total, 0) + ISNULL(Devolucoes.ValorFuturo, 0))) as ValorLiquido,"
                SqlT &= "         CONCAT('REF. NF ',Nota.Nota_Id,' - ',Nota.Serie_Id,' - ',Nota.EntradaSaida_Id,' DE ',ProdutoNota.Nome,'- Pedido ',PedidoNota.Pedido_Id) as Historico,"
                SqlT &= "          '' as Observacoes,"
                SqlT &= "          '' as Destinatario,"
                SqlT &= "          '' as NomeDoDestinatario,"
                SqlT &= "          ProdutoNota.Produto_Id as Produto,"
                SqlT &= "          ProdutoNota.Nome as NomeProduto,"
                SqlT &= "   (SELECT SUM(EncargoLiquido.Valor) FROM  NotasFiscaisXItens AS ItensNota     "
                SqlT &= "    INNER Join NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNota.Empresa_Id    "
                SqlT &= "    And EncargoLiquido.EndEmpresa_Id = ItensNota.EndEmpresa_Id     "
                SqlT &= "    And EncargoLiquido.Cliente_Id = ItensNota.Cliente_Id     "
                SqlT &= "    And EncargoLiquido.EndCliente_Id = ItensNota.EndCliente_Id     "
                SqlT &= "    And EncargoLiquido.EntradaSaida_Id = ItensNota.EntradaSaida_Id    "
                SqlT &= "    And EncargoLiquido.Serie_Id = ItensNota.Serie_Id   "
                SqlT &= "    And EncargoLiquido.Nota_Id = ItensNota.Nota_Id    "
                SqlT &= "    And EncargoLiquido.Produto_Id = ItensNota.Produto_Id     "
                SqlT &= "    And EncargoLiquido.Encargo_Id = 'LIQUIDO'     "
                SqlT &= "    WHERE ItensNota.Empresa_Id = Nota.Empresa_Id     "
                SqlT &= "    And ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id     "
                SqlT &= "    And ItensNota.Cliente_Id = Nota.Cliente_Id    "
                SqlT &= "    And ItensNota.EndCliente_Id = Nota.EndCliente_Id     "
                SqlT &= "    And ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id    "
                SqlT &= "    And ItensNota.Serie_Id = Nota.Serie_Id    "
                SqlT &= "    And ItensNota.Nota_Id = Nota.Nota_Id     "
                SqlT &= "   ) AS Total,     "
                SqlT &= "       (SELECT ISNULL(SUM(CASE"
                SqlT &= "                            WHEN TipoDeLancamento = 'E'"
                SqlT &= "                              THEN ISNULL(pXi.TotalMoeda, 0) * - 1"
                SqlT &= "                              ELSE ISNULL(pXi.TotalMoeda, 0)"
                SqlT &= "                          END),0)"
                SqlT &= "          FROM Pedidos"
                SqlT &= "          LEFT JOIN PedidoXItemxLancamento AS pXi"
                SqlT &= "            ON Pedidos.Pedido_Id     = pXi.Pedido_Id"
                SqlT &= "           AND Pedidos.Empresa_Id    = pXi.Empresa_Id"
                SqlT &= "           AND Pedidos.EndEmpresa_Id = pXi.EndEmpresa_Id	"
                SqlT &= "         WHERE Pedidos.Pedido_Id     = PedidoNota.Pedido_Id"
                SqlT &= "           AND Pedidos.Empresa_Id    = PedidoNota.Empresa_Id"
                SqlT &= "           AND Pedidos.EndEmpresa_Id = PedidoNota.EndEmpresa_Id) AS TotalEmDolar,"
                SqlT &= "           Nota.Nota_Id as Nota,"
                SqlT &= "           CONVERT(VARCHAR,Nota.Movimento,103) as MovimentoNota,"
                SqlT &= "           '' as ContratoBanco,"
                SqlT &= "           '' as RegistroMestre,"
                SqlT &= "           SoPedido.Classe,"
                SqlT &= "           '' as ContraPartida,"
                SqlT &= "           '' as DescricaoContrapartida,"
                SqlT &= "           '' as TipoDaConta,"
                SqlT &= "	   isnull((SELECT top 1 raz.Conta_id "
                SqlT &= "	      from razao raz "
                SqlT &= "			where raz.empresa_id    = Nota.Empresa_Id "
                SqlT &= "			and raz.endempresa_id   = Nota.EndEmpresa_Id "
                SqlT &= "			and raz.cliente_nf      = Nota.Cliente_Id "
                SqlT &= "			and raz.endcliente_nf   = Nota.EndCliente_Id "
                SqlT &= "			and raz.entradasaida_nf = Nota.EntradaSaida_Id "
                SqlT &= "			and raz.serie_nf        = Nota.Serie_Id "
                SqlT &= "			and raz.numero_nf       = Nota.Nota_Id "
                SqlT &= "			and raz.Encargo_NF = 'PRODUTO'),'') AS ContrapartidaNota,"
                SqlT &= "	   isnull((SELECT top 1 raz.Conta_id as Titulo "
                SqlT &= "	      from razao raz "
                SqlT &= "			where raz.empresa_id    = Nota.Empresa_Id "
                SqlT &= "			and raz.endempresa_id   = Nota.EndEmpresa_Id "
                SqlT &= "			and raz.cliente_nf      = Nota.Cliente_Id"
                SqlT &= "			and raz.endcliente_nf   = Nota.EndCliente_Id "
                SqlT &= "			and raz.entradasaida_nf = Nota.EntradaSaida_Id "
                SqlT &= "			and raz.serie_nf        = Nota.Serie_Id "
                SqlT &= "			and raz.numero_nf       = Nota.Nota_Id "
                SqlT &= "			and raz.Encargo_NF = 'PRODUTO'),'') AS DescricaoContrapartidaNota,'' Tipo"

            Else

                SqlT &= "  SELECT  Empresa.Reduzido AS ReduzidoEmpresa,"
                SqlT &= "          Empresa.Cliente_Id as Empresa,"
                SqlT &= "          Empresa.Endereco_Id as EndEmpresa,"
                SqlT &= "          Empresa.Nome AS NomeEmpresa,"
                SqlT &= "          Empresa.Cidade AS CidadeEmpresa,"
                SqlT &= "          Empresa.Estado AS EstadoEmpresa,"
                SqlT &= "          convert(varchar,Nota.Nota_Id) AS Titulo,"
                SqlT &= "          convert(nvarchar,Nota.Nota_Id) + ' ' + CASE"
                SqlT &= "														 WHEN isnull(PedidoNota.Moeda, 1) = 1 THEN '(NOTA) R$'"
                SqlT &= "													     ELSE '(NOTA) U$'"
                SqlT &= "                                                 END Registro,"
                SqlT &= "          '' AS Faturamento,"
                SqlT &= "          '' AS Lote,"
                SqlT &= "          0 AS LoteTotal,"
                SqlT &= "          0 AS LoteEntregue,"
                SqlT &= "          PedidoNota.Pedido_Id as Pedido,"
                SqlT &= "          Cliente.Cliente_Id as Cliente,"
                SqlT &= "          Cliente.Nome AS NomeCliente,"
                SqlT &= "          Nota.Movimento,"
                SqlT &= "          Nota.Movimento AS Vencimento,"
                SqlT &= "          Nota.Movimento AS Baixa,"
                SqlT &= "          PedidoNota.Carteira,"
                SqlT &= "          CarteiraPedido.Descricao AS NomeCarteira,"
                SqlT &= "          CONCAT('REF. NF ',Nota.Nota_Id,' - ',Nota.Serie_Id,' - ',Nota.EntradaSaida_Id,' DE ',ProdutoNota.Nome,'- Pedido ',PedidoNota.Pedido_Id) as Historico,"
                SqlT &= "          '' as solicitacao,"
                SqlT &= "          ISNULL(PedidoNota.PedidoEfetivo, 0) PedidoEfetivo,"
                SqlT &= "          '' AS Provisao,"
                SqlT &= "         ((SUM(ISNULL(EncargoLiquidoNotaProduto.Valor, 0)) - ISNULL(Devolucoes.ValorPeriodo, 0)) - (ISNULL(TitulosBaixados.Total, 0) + ISNULL(Devolucoes.ValorFuturo, 0))) AS ValorLiquido,"
                SqlT &= "          0 AS MoedaValorLiquido,"
                SqlT &= "          '' AS tipo"

            End If

            SqlT &= "   FROM NotasFiscais as Nota "
            SqlT &= ""
            SqlT &= "   INNER JOIN Clientes AS Empresa ON Empresa.Cliente_Id  = Nota.Empresa_Id"
            SqlT &= "   AND Empresa.Endereco_Id = Nota.EndEmpresa_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN SubOperacoes AS SoNota ON SoNota.Operacao_Id  = Nota.Operacao"
            SqlT &= "   AND SoNota.SubOperacoes_Id = Nota.SubOperacao"
            SqlT &= "   AND SoNota.Financeiro   = 'S'"
            SqlT &= "   AND SoNota.Devolucao    = 'N'"
            SqlT &= ""
            SqlT &= "   INNER JOIN Clientes AS Cliente ON Cliente.Cliente_Id  = Nota.Cliente_Id"
            SqlT &= "   AND Cliente.Endereco_Id = Nota.EndCliente_Id"
            SqlT &= ""
            SqlT &= "   INNER JOIN Pedidos as PedidoNota on PedidoNota.Empresa_Id = Nota.Empresa_Id"
            SqlT &= "   AND PedidoNota.EndEmpresa_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND PedidoNota.Cliente = Nota.Cliente_Id"
            SqlT &= "   AND PedidoNota.EndCliente = Nota.EndCliente_Id"
            SqlT &= "   AND PedidoNota.Pedido_Id = Nota.Pedido"
            SqlT &= ""
            SqlT &= "   INNER JOIN SubOperacoes AS SoPedido ON SoPedido.Operacao_Id  = PedidoNota.Operacao"
            SqlT &= "   AND SoPedido.SubOperacoes_Id = PedidoNota.SubOperacao"
            SqlT &= ""
            SqlT &= "   INNER JOIN Carteira as CarteiraPedido on CarteiraPedido.Carteira_Id = PedidoNota.Carteira"
            SqlT &= ""
            SqlT &= "   INNER JOIN NotasFiscaisXItens AS ItensNota ON ItensNota.Empresa_Id = Nota.Empresa_Id"
            SqlT &= "   AND ItensNota.EndEmpresa_Id = Nota.EndEmpresa_Id"
            SqlT &= "   AND ItensNota.Cliente_Id = Nota.Cliente_Id"
            SqlT &= "   AND ItensNota.EndCliente_Id = Nota.EndCliente_Id"
            SqlT &= "   AND ItensNota.EntradaSaida_Id = Nota.EntradaSaida_Id"
            SqlT &= "   AND ItensNota.Serie_Id = Nota.Serie_Id"
            SqlT &= "   AND ItensNota.Nota_Id = Nota.Nota_Id"
            SqlT &= ""
            SqlT &= "   Left Join NotasFiscaisXEncargos AS EncargoLiquidoNotaProduto ON EncargoLiquidoNotaProduto.Empresa_Id = ItensNota.Empresa_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.EndEmpresa_Id = ItensNota.EndEmpresa_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.Cliente_Id = ItensNota.Cliente_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.EndCliente_Id = ItensNota.EndCliente_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.EntradaSaida_Id = ItensNota.EntradaSaida_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.Serie_Id = ItensNota.Serie_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.Nota_Id = ItensNota.Nota_Id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.Sequencia_id = ItensNota.Sequencia_id"
            SqlT &= "   AND EncargoLiquidoNotaProduto.Produto_Id = ItensNota.Produto_Id"
            SqlT &= "   And EncargoLiquidoNotaProduto.Encargo_Id = 'LIQUIDO'"

            SqlT &= "   OUTER APPLY(SELECT SUM(Contas.ValorDoDocumento) as Total " & vbCrLf &
                    "               FROM NotaFiscalXTitulo As Titulo " & vbCrLf


            If RdReceber.Checked Then
                SqlT &= " INNER Join ContasAReceber AS Contas ON Contas.Registro_Id = Titulo.Titulo_Id AND Contas.Situacao = '1' " & vbCrLf
            Else
                SqlT &= "INNER Join ContasAPagar AS Contas ON Contas.Registro_Id = Titulo.Titulo_Id AND Contas.Situacao = '1' " & vbCrLf
            End If

            SqlT &= " WHERE Titulo.Empresa_Id = Nota.Empresa_Id " & vbCrLf &
                    " AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id " & vbCrLf &
            "         AND   Titulo.Cliente_Id = Nota.Cliente_Id" & vbCrLf &
            "         AND   Titulo.EndCliente_Id = Nota.EndCliente_Id " & vbCrLf &
            "         AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id " & vbCrLf &
            "         AND   Titulo.Serie_Id = Nota.Serie_Id " & vbCrLf &
            "         AND   Titulo.Nota_Id = Nota.Nota_Id " & vbCrLf &
            "         AND   Contas.Provisao IN (1,2) ) TitulosBaixados" & vbCrLf

            SqlT &= "   INNER JOIN Produtos AS ProdutoNota ON ProdutoNota.Produto_Id = ItensNota.Produto_Id"

            SqlT &= ""
            SqlT &= " OUTER APPLY ("
            SqlT &= "   SELECT SUM(ISNULL(EncargoLiquido.Valor, 0)) AS ValorFuturo,"
            SqlT &= "          SUM(ISNULL(EncargoLiquidoPeriodo.Valor, 0)) AS ValorPeriodo"
            SqlT &= "    FROM NotaFiscalDevolucaoXNotaFiscal AS Devolucoes "
            SqlT &= "    LEFT JOIN NotasFiscais AS NotaDevolucao ON NotaDevolucao.Empresa_Id = Devolucoes.EmpresaDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.EndEmpresa_Id = Devolucoes.EndEmpresaDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.Cliente_Id = Devolucoes.ClienteDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.EndCliente_Id = Devolucoes.EndClienteDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.EntradaSaida_Id = Devolucoes.EntradaSaidaDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.Serie_Id = Devolucoes.SerieDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.Nota_Id = Devolucoes.NotaDevolucao_Id "
            SqlT &= "    AND NotaDevolucao.Movimento > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
            SqlT &= "    LEFT JOIN NotasFiscaisXItens AS ItensNotaDevolucao ON ItensNotaDevolucao.Empresa_Id = NotaDevolucao.Empresa_Id "
            SqlT &= "    AND ItensNotaDevolucao.EndEmpresa_Id = NotaDevolucao.EndEmpresa_Id "
            SqlT &= "    AND ItensNotaDevolucao.Cliente_Id = NotaDevolucao.Cliente_Id "
            SqlT &= "    AND ItensNotaDevolucao.EndCliente_Id = NotaDevolucao.EndCliente_Id "
            SqlT &= "    AND ItensNotaDevolucao.EntradaSaida_Id = NotaDevolucao.EntradaSaida_Id "
            SqlT &= "    AND ItensNotaDevolucao.Serie_Id = NotaDevolucao.Serie_Id "
            SqlT &= "    AND ItensNotaDevolucao.Nota_Id = NotaDevolucao.Nota_Id "
            SqlT &= "    AND ItensNotaDevolucao.CFOP_Id = Devolucoes.CFOPDevolucao_Id "
            SqlT &= "    AND ItensNotaDevolucao.Produto_Id = Devolucoes.Produto_Id "
            SqlT &= "    LEFT JOIN NotasFiscaisXEncargos AS EncargoLiquido ON EncargoLiquido.Empresa_Id = ItensNotaDevolucao.Empresa_Id "
            SqlT &= "    AND EncargoLiquido.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id "
            SqlT &= "    AND EncargoLiquido.Cliente_Id = ItensNotaDevolucao.Cliente_Id "
            SqlT &= "    AND EncargoLiquido.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id "
            SqlT &= "    AND EncargoLiquido.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id "
            SqlT &= "    AND EncargoLiquido.Serie_Id = ItensNotaDevolucao.Serie_Id "
            SqlT &= "    AND EncargoLiquido.Nota_Id = ItensNotaDevolucao.Nota_Id "
            SqlT &= "    AND EncargoLiquido.Sequencia_id = ItensNotaDevolucao.Sequencia_id "
            SqlT &= "    AND EncargoLiquido.Produto_Id = ItensNotaDevolucao.Produto_Id "
            SqlT &= "    AND EncargoLiquido.Encargo_Id = 'LIQUIDO' "
            SqlT &= "    LEFT JOIN NotasFiscaisXEncargos AS EncargoProduto ON EncargoProduto.Empresa_Id = ItensNotaDevolucao.Empresa_Id "
            SqlT &= "    AND EncargoProduto.EndEmpresa_Id = ItensNotaDevolucao.EndEmpresa_Id "
            SqlT &= "    AND EncargoProduto.Cliente_Id = ItensNotaDevolucao.Cliente_Id "
            SqlT &= "    AND EncargoProduto.EndCliente_Id = ItensNotaDevolucao.EndCliente_Id "
            SqlT &= "    AND EncargoProduto.EntradaSaida_Id = ItensNotaDevolucao.EntradaSaida_Id "
            SqlT &= "    AND EncargoProduto.Serie_Id = ItensNotaDevolucao.Serie_Id "
            SqlT &= "    AND EncargoProduto.Nota_Id = ItensNotaDevolucao.Nota_Id "
            SqlT &= "    AND EncargoProduto.Sequencia_id = ItensNotaDevolucao.Sequencia_id "
            SqlT &= "    AND EncargoProduto.Produto_Id = ItensNotaDevolucao.Produto_Id "
            SqlT &= "    AND EncargoProduto.Encargo_Id = 'PRODUTO' "
            SqlT &= "    Left Join NotasFiscais AS NotaDevolucaoPeriodo ON NotaDevolucaoPeriodo.Empresa_Id = Devolucoes.EmpresaDevolucao_Id "
            SqlT &= "    And NotaDevolucaoPeriodo.EndEmpresa_Id = Devolucoes.EndEmpresaDevolucao_Id"
            SqlT &= "    And NotaDevolucaoPeriodo.Cliente_Id = Devolucoes.ClienteDevolucao_Id"
            SqlT &= "    And NotaDevolucaoPeriodo.EndCliente_Id = Devolucoes.EndClienteDevolucao_Id"
            SqlT &= "    And NotaDevolucaoPeriodo.EntradaSaida_Id = Devolucoes.EntradaSaidaDevolucao_Id"
            SqlT &= "    And NotaDevolucaoPeriodo.Serie_Id = Devolucoes.SerieDevolucao_Id"
            SqlT &= "    And NotaDevolucaoPeriodo.Nota_Id = Devolucoes.NotaDevolucao_Id"
            SqlT &= "    And NotaDevolucaoPeriodo.Movimento <= '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
            SqlT &= "    Left Join NotasFiscaisXItens AS ItensNotaDevolucaoPeriodo ON ItensNotaDevolucaoPeriodo.Empresa_Id = NotaDevolucaoPeriodo.Empresa_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.EndEmpresa_Id = NotaDevolucaoPeriodo.EndEmpresa_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.Cliente_Id = NotaDevolucaoPeriodo.Cliente_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.EndCliente_Id = NotaDevolucaoPeriodo.EndCliente_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.EntradaSaida_Id = NotaDevolucaoPeriodo.EntradaSaida_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.Serie_Id = NotaDevolucaoPeriodo.Serie_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.Nota_Id = NotaDevolucaoPeriodo.Nota_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.CFOP_Id = Devolucoes.CFOPDevolucao_Id"
            SqlT &= "    And ItensNotaDevolucaoPeriodo.Produto_Id = Devolucoes.Produto_Id"
            SqlT &= "    Left Join NotasFiscaisXEncargos AS EncargoLiquidoPeriodo ON EncargoLiquidoPeriodo.Empresa_Id = ItensNotaDevolucaoPeriodo.Empresa_Id"
            SqlT &= "    And EncargoLiquidoPeriodo.EndEmpresa_Id = ItensNotaDevolucaoPeriodo.EndEmpresa_Id"
            SqlT &= "    And EncargoLiquidoPeriodo.Cliente_Id = ItensNotaDevolucaoPeriodo.Cliente_Id "
            SqlT &= "    And EncargoLiquidoPeriodo.EndCliente_Id = ItensNotaDevolucaoPeriodo.EndCliente_Id "
            SqlT &= "    And EncargoLiquidoPeriodo.EntradaSaida_Id = ItensNotaDevolucaoPeriodo.EntradaSaida_Id "
            SqlT &= "    And EncargoLiquidoPeriodo.Serie_Id = ItensNotaDevolucaoPeriodo.Serie_Id"
            SqlT &= "    And EncargoLiquidoPeriodo.Nota_Id = ItensNotaDevolucaoPeriodo.Nota_Id"
            SqlT &= "    And EncargoLiquidoPeriodo.Sequencia_id = ItensNotaDevolucaoPeriodo.Sequencia_id"
            SqlT &= "    And EncargoLiquidoPeriodo.Produto_Id = ItensNotaDevolucaoPeriodo.Produto_Id"
            SqlT &= "    And EncargoLiquidoPeriodo.Encargo_Id = 'LIQUIDO'"
            SqlT &= "    WHERE Devolucoes.EmpresaDevolucao_Id = Nota.Empresa_Id "
            SqlT &= "    AND   Devolucoes.EndEmpresaDevolucao_Id = Nota.EndEmpresa_Id "
            SqlT &= "    AND   Devolucoes.ClienteDevolucao_Id = Nota.Cliente_Id "
            SqlT &= "    AND   Devolucoes.EndClienteDevolucao_Id = Nota.EndCliente_Id "
            SqlT &= "    AND   Devolucoes.Produto_Id = ItensNota.Produto_Id "
            SqlT &= "    AND   Devolucoes.EntradaSaida_Id = Nota.EntradaSaida_Id "
            SqlT &= "    AND   Devolucoes.Serie_Id = Nota.Serie_Id "
            SqlT &= "    AND   Devolucoes.Nota_Id = Nota.Nota_Id "
            SqlT &= "   ) AS Devolucoes  "


            If ddlSafra.SelectedIndex > 0 Then
                SqlT &= "  Where PedidoNota.Safra = '" & ddlSafra.SelectedValue & "' " & vbCrLf
            Else
                SqlT &= "  Where 1 = 1"
            End If

            If op.Length > 0 Then
                SqlT &= "   AND isnull(PedidoNota.Finalidade,0) " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & op & "')"
            End If

            If ddlRepresentante.SelectedIndex > 0 Then
                SqlT &= "  And exists (Select 1" & vbCrLf &
                        "                from Comissoes C" & vbCrLf &
                        "               where C.Empresa_id          = PedidoNota.Empresa_id" & vbCrLf &
                        "                 And C.EndEmpresa_Id       = PedidoNota.EndEmpresa_Id " & vbCrLf &
                        "                 And C.Pedido_id           = PedidoNota.Pedido_id " & vbCrLf &
                        "                 And C.Representante_id    ='" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                        "                 And C.EndRepresentante_Id = " & ddlRepresentante.SelectedValue.Split("-")(1) & ")" & vbCrLf
            End If

            SqlT &= "    AND Nota.Situacao = 1 "

            'Comentei por não estar aparecendo os Titulos Lançados Tirando do Banco para o Caixa e de Transferência entre Contas - 31/03/2016 - Furlan
            '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf & _
            '"    AND (Titulos.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf


            SqlT &= " AND Nota.movimento BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf

            If ddlMoeda.SelectedIndex > 0 Then
                SqlT &= " and PedidoNota.Moeda = " & ddlMoeda.SelectedValue & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    SqlT &= " and Empresa.Cliente_Id in(" & EmpresasConsolidas & ")" & vbCrLf
                Else
                    SqlT &= " and Empresa.Cliente_Id = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            " and Empresa.Endereco_Id = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
                End If
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                If chkConsolidarCliente.Checked Then
                    SqlT &= " and left(Cliente.Cliente_Id,8) = '" & txtCodigoCliente.Value.Split("-")(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    SqlT &= " and Cliente.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                            " and Cliente.Endereco_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
                End If
            End If

            If ddlCarteiraDoTitulo.SelectedIndex > 0 Then
                SqlT &= " and PedidoNota.Carteira = '" & ddlCarteiraDoTitulo.SelectedValue & "'" & vbCrLf
            End If

            If RdReceber.Checked Then

                SqlT &= " AND NOT EXISTS(Select top 1 * from NotaFiscalXTitulo as Titulo "
                SqlT &= ""
                SqlT &= "			     INNER JOIN ContasAReceber ON ContasAReceber.Registro_Id = Titulo.Titulo_Id"
                SqlT &= ""
                SqlT &= "                WHERE Titulo.Empresa_Id = Nota.Empresa_Id"
                SqlT &= "			     AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id"
                SqlT &= "			     AND   Titulo.Cliente_Id = Nota.Cliente_Id"
                SqlT &= "			     AND   Titulo.EndCliente_Id = Nota.EndCliente_Id"
                SqlT &= "			     AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id"
                SqlT &= "		         AND   Titulo.Serie_Id = Nota.Serie_Id"
                SqlT &= "		         AND   Titulo.Nota_Id = Nota.Nota_Id "
                SqlT &= "                AND ContasAReceber.Provisao = 2)"

                SqlT &= "AND Nota.EntradaSaida_Id = 'S' "
            ElseIf RdPagar.Checked Then

                SqlT &= " AND NOT EXISTS(Select top 1 * from NotaFiscalXTitulo as Titulo "
                SqlT &= ""
                SqlT &= "                INNER JOIN ContasAPagar ON ContasAPagar.Registro_Id = Titulo.Titulo_Id"
                SqlT &= ""
                SqlT &= "                WHERE Titulo.Empresa_Id = Nota.Empresa_Id"
                SqlT &= "			     AND   Titulo.EndEmpresa_Id = Nota.EndEmpresa_Id"
                SqlT &= "			     AND   Titulo.Cliente_Id = Nota.Cliente_Id"
                SqlT &= "			     AND   Titulo.EndCliente_Id = Nota.EndCliente_Id"
                SqlT &= "			     AND   Titulo.EntradaSaida_Id = Nota.EntradaSaida_Id"
                SqlT &= "		         AND   Titulo.Serie_Id = Nota.Serie_Id"
                SqlT &= "		         AND   Titulo.Nota_Id = Nota.Nota_Id "
                SqlT &= "                AND   ContasAPagar.Provisao = 2)"

                SqlT &= "AND Nota.EntradaSaida_Id = 'E' "

            End If

            If Excel Then

                SqlT &= ""
                SqlT &= " GROUP BY   Nota.Nota_id,"
                SqlT &= "            Nota.Movimento,"
                SqlT &= "            Empresa.Reduzido,"
                SqlT &= "            Empresa.Nome,"
                SqlT &= "            Empresa.Cidade,"
                SqlT &= "            Cliente.Reduzido,"
                SqlT &= "            Cliente.Nome,"
                SqlT &= "            PedidoNota.Pedido_Id,"
                SqlT &= "            Nota.Serie_Id,"
                SqlT &= "            Nota.EntradaSaida_Id,"
                SqlT &= "            ProdutoNota.Produto_Id,"
                SqlT &= "            ProdutoNota.Nome,"
                SqlT &= "            SoPedido.Classe,"
                SqlT &= "            PedidoNota.Empresa_Id,"
                SqlT &= "            PedidoNota.EndEmpresa_Id,"
                SqlT &= "            Nota.Empresa_Id,"
                SqlT &= "            Nota.EndEmpresa_Id,"
                SqlT &= "            Nota.Cliente_Id,"
                SqlT &= "            Nota.EndCliente_Id,"
                SqlT &= "            TitulosBaixados.Total, "
                SqlT &= "            Devolucoes.ValorFuturo, "
                SqlT &= "            Devolucoes.ValorPeriodo "
            Else

                SqlT &= ""
                SqlT &= " GROUP BY   Nota.Nota_id,"
                SqlT &= "		     Empresa.Reduzido,"
                SqlT &= "		     Empresa.Cliente_Id,"
                SqlT &= "		     Empresa.Endereco_Id,"
                SqlT &= "		     Empresa.Nome,"
                SqlT &= "		     Empresa.Cidade,"
                SqlT &= "		     Empresa.Estado,"
                SqlT &= "		     PedidoNota.Moeda, "
                SqlT &= "	         PedidoNota.Pedido_Id,"
                SqlT &= "		     Cliente.Cliente_Id, "
                SqlT &= "	         Cliente.Nome,"
                SqlT &= "            Nota.Movimento,"
                SqlT &= "            PedidoNota.Carteira,"
                SqlT &= "            CarteiraPedido.Descricao,"
                SqlT &= "            Nota.Serie_Id,"
                SqlT &= "            Nota.EntradaSaida_Id,"
                SqlT &= "            ProdutoNota.Nome,"
                SqlT &= "            PedidoNota.PedidoEfetivo,"
                SqlT &= "            TitulosBaixados.Total, "
                SqlT &= "            Devolucoes.ValorFuturo, "
                SqlT &= "            Devolucoes.ValorPeriodo "

            End If



            SqlT &= ""
            SqlT &= "  HAVING ((SUM(ISNULL(EncargoLiquidoNotaProduto.Valor, 0)) - ISNULL(Devolucoes.ValorPeriodo, 0)) - (ISNULL(TitulosBaixados.Total, 0) + ISNULL(Devolucoes.ValorFuturo, 0))) > 0 "

        End If



        Return SqlT
    End Function

    Private Function GetProvisao(DataBase As Boolean) As String
        Dim Provisao As String = String.Empty

        If DataBase Then
            Provisao = "       Case  " & vbCrLf &
                       "         When Titulos.baixa > '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "' and Titulos.Provisao = 1 " & vbCrLf &
                       "           Then 2 " & vbCrLf &
                       "           Else Titulos.Provisao " & vbCrLf &
                       "        end"
        Else
            Provisao = "       Titulos.Provisao"
        End If

        Return Provisao

    End Function

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Function GetSQLTituloNovo() As String
        Dim sql As String = String.Empty
        Dim parametros As String = String.Empty

        sql = " SELECT Empresa.Reduzido As ReduzidoEmpresa, " & vbCrLf &
              "        Empresa.Cliente_Id AS Empresa, " & vbCrLf &
              "        Empresa.Endereco_Id AS EndEmpresa, " & vbCrLf &
              "        Empresa.Nome AS NomeEmpresa, " & vbCrLf &
              "        Empresa.Cidade AS CidadeEmpresa, " & vbCrLf &
              "        Empresa.Estado AS EstadoEmpresa, " & vbCrLf &
              "        CONVERT(NVARCHAR, T.Titulo_Id) + ' ' + " & vbCrLf &
              "           CASE  " & vbCrLf &
              "             WHEN ISNULL(T.Moeda,1) = 1 " & vbCrLf &
              "             THEN 'R$' " & vbCrLf &
              "               ELSE 'U$' " & vbCrLf &
              "            END AS Registro, " & vbCrLf &
              "        ISNULL(T.Pedido, 0) AS Pedido, " & vbCrLf &
              "        Cliente.Cliente_Id AS Cliente, " & vbCrLf &
              "        Cliente.Nome AS NomeCliente, " & vbCrLf &
              "        T.Movimento, " & vbCrLf &
              "        T.Reprogramacao AS Vencimento, " & vbCrLf &
              "        T.DataBaixa AS Baixa, " & vbCrLf &
              "        T.Provisao, " & vbCrLf &
              "        T.CarteiraDoTitulo AS Carteira, " & vbCrLf &
              "        Carteira.Descricao AS NomeCarteira, " & vbCrLf &
              "        T.Historico, " & vbCrLf &
              "        0 AS Solicitacao, " & vbCrLf &
              "        CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf &
              "            (SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C') " & vbCrLf &
              "        ELSE " & vbCrLf &
              "            (SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D') " & vbCrLf &
              "        END AS ValorLiquido, " & vbCrLf &
              "        CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf &
              "            (SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C') " & vbCrLf &
              "        ELSE " & vbCrLf &
              "            (SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D') " & vbCrLf &
              "        END AS MoedaValorLiquido, " & vbCrLf &
              "        tf.Descricao as Faturamento, " & vbCrLf &
              "        (SELECT  " & vbCrLf &
              "           case when tit.Situacao = 3 and p.MomentoFinanceiro = 4 then 'Encerrado' " & vbCrLf &
              "                when p.MomentoFinanceiro = 4 and tit.Situacao <> 3 then 'Em andamento' end as Posicao " & vbCrLf &
              "     FROM Titulos AS tit " & vbCrLf &
              " left join Pedidos as p " & vbCrLf &
              "    on p.Pedido_Id = tit.Pedido " & vbCrLf &
              " left join TipoDeFaturamento as tf " & vbCrLf &
              "    on p.MomentoFinanceiro = tf.TipoDeFaturamento_Id " & vbCrLf &
              " WHERE(tit.Titulo_Id = t.tituloOrigem) " & vbCrLf &
              " ) AS Lote, " & vbCrLf

        'Quantidade Total do Lote
        sql &= "    ( SELECT  " & vbCrLf &
               "        SUM(tit.Quantidade) " & vbCrLf &
               "      FROM Titulos AS tit  " & vbCrLf &
               "      WHERE tit.Titulo_Id = T.tituloOrigem) AS LoteTotal, " & vbCrLf

        'Quantidade Total entregue do Lote
        sql &= "    ( SELECT  " & vbCrLf &
               "        SUM(tit.Quantidade) " & vbCrLf &
               "      FROM  Titulos AS tit  " & vbCrLf &
               "      WHERE tit.TituloOrigem = T.tituloOrigem) AS LoteEntregue " & vbCrLf

        sql &= " FROM Titulos AS T " & vbCrLf &
               "    INNER JOIN Clientes AS Empresa " & vbCrLf &
               "        ON T.Empresa = Empresa.Cliente_Id " & vbCrLf &
               "        AND T.EndEmpresa = Empresa.Endereco_Id " & vbCrLf &
               "    INNER JOIN Clientes AS Cliente " & vbCrLf &
               "        ON T.CliFor = Cliente.Cliente_Id " & vbCrLf &
               "        AND T.EnderecoCliFor = Cliente.Endereco_Id	" & vbCrLf &
               "    LEFT OUTER JOIN Carteira " & vbCrLf &
               "         ON T.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf &
               "    left join Pedidos as p " & vbCrLf &
               "       on p.Pedido_Id = t.Pedido " & vbCrLf &
               "    left join TipoDeFaturamento as tf " & vbCrLf &
               "       on p.MomentoFinanceiro = tf.TipoDeFaturamento_Id  " & vbCrLf

        If (Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Or Not String.IsNullOrWhiteSpace(ddlRepresentante.SelectedValue)) Then
            sql &= "INNER JOIN Pedidos AS P	" & vbCrLf &
                   "    ON P.Empresa_Id = T.Empresa " & vbCrLf &
                   "    AND P.EndEmpresa_Id = T.EndEmpresa " & vbCrLf &
                   "    AND P.Pedido_Id = T.Pedido " & vbCrLf
        End If

        If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
            sql &= " WHERE T.Situacao IN(1,101,102,105)" & vbCrLf
        Else
            sql &= " WHERE T.Situacao IN(1,101,102)" & vbCrLf
        End If

        If (Not String.IsNullOrEmpty(DdlUnidadeConsultaTitulos.SelectedValue)) Then
            sql &= " AND T.UnidadeDeNegocio = " & DdlUnidadeConsultaTitulos.SelectedValue & vbCrLf
        End If

        If Not String.IsNullOrEmpty(DdlEmpresaConsultaTitulos.SelectedValue) Then
            If chkConsolidarEmpresa.Checked Then
                sql &= " AND LEFT(T.Empresa, 8) = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0).Substring(0, 8) & vbCrLf
            Else
                sql &= " AND T.Empresa = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                       " AND T.EndEmpresa = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
            End If
        End If

        Dim Cliente = txtCodigoCliente.Value.Split("-")
        If (Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value)) Then
            If (chkConsolidarCliente.Checked) Then
                sql &= " AND LEFT(T.CliFor, 8) = " & Cliente(0).Substring(0, 8) & vbCrLf
            Else
                sql &= " AND T.CliFor = '" & Cliente(0) & "'" & vbCrLf &
                       " AND T.EnderecoCliFor = " & Cliente(1) & vbCrLf
            End If
        End If

        If (Not String.IsNullOrWhiteSpace(ddlRepresentante.SelectedValue)) Then
            sql &= "  And exists (Select 1" & vbCrLf &
                    "                from Comissoes C" & vbCrLf &
                    "               where C.Empresa_id          = P.Empresa_id" & vbCrLf &
                    "                 And C.EndEmpresa_Id       = P.EndEmpresa_Id " & vbCrLf &
                    "                 And C.Pedido_id           = P.Pedido_id " & vbCrLf &
                    "                 And P.Sequencia_id        = 0" & vbCrLf &
                    "                 And C.Representante_id    ='" & ddlRepresentante.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                    "                 And C.EndRepresentante_Id = " & ddlRepresentante.SelectedValue.Split("-")(1) & ")" & vbCrLf

            parametros &= "Representante: " & ddlRepresentante.SelectedItem.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlCarteiraDoTitulo.SelectedValue) Then
            sql &= " AND T.CarteiraDoTitulo = '" & ddlCarteiraDoTitulo.SelectedValue & "'" & vbCrLf
            parametros &= "Carteira: " & ddlCarteiraDoTitulo.SelectedItem.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            sql &= "  AND P.Safra         = '" & ddlSafra.SelectedValue & "' " & vbCrLf
            parametros &= "safra: " & ddlSafra.SelectedValue & vbCrLf
        End If

        If IsDate(txtPeriodoInicialConsultaTitulos.Text) And IsDate(txtPeriodoFinalConsultaTitulos.Text) And chkPeriodo.Checked Then
            sql &= " AND T.Reprogramacao BETWEEN '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' AND '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf
            parametros &= "Vecimentos entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text
        End If

        If (Not RdReceberPagar.Checked) Then
            sql &= " AND RecPag = " & IIf(RdPagar.Checked, "'P'", "'R'") & vbCrLf
        End If

        If RadAtivos.Checked Then
            If chkPrevisao.Visible And chkProvisao.Visible Then
                If chkPrevisao.Checked And chkProvisao.Checked Then
                    sql &= " AND T.Provisao IN(2,3) " & vbCrLf
                    parametros &= "Titulos Previsionados, Provisionados" & vbCrLf
                ElseIf chkPrevisao.Checked Then
                    sql &= " AND T.Provisao = 2 " & vbCrLf
                    parametros &= "Titulos Previsionados" & vbCrLf
                ElseIf chkProvisao.Checked Then
                    sql &= " AND T.Provisao = 3 " & vbCrLf
                    parametros &= "Titulos Provisionados" & vbCrLf
                Else
                    sql &= " AND T.Provisao = 2 " & vbCrLf
                    parametros &= "Titulos Previsionados" & vbCrLf
                End If
            Else
                sql &= " AND T.Provisao = 2 " & vbCrLf
                parametros &= "Titulos Previsionados" & vbCrLf
            End If
            parametros &= "Titulos Abertos" & vbCrLf
        End If

        If RadBaixados.Checked Then
            sql &= " AND T.Provisao = 1 " & vbCrLf
            parametros &= "Titulos Baixados" & vbCrLf
        End If

        If ddlMoeda.SelectedIndex > 0 Then
            sql &= " AND T.Moeda = " & ddlMoeda.SelectedValue & vbCrLf
            parametros &= "Moeda: " & ddlMoeda.SelectedItem.Text & vbCrLf
        End If
        Return sql
    End Function

    Protected Sub lnkPlanilhaExcel_Click(sender As Object, e As EventArgs) Handles lnkPlanilhaExcel.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioDeTitulos", "RELATORIO") Then
                If chkMovtoBancario.Checked Then
                    Dim ds As DataSet = getDataSetMovtoBancario()
                    Dim colunas As New Dictionary(Of String, eTipoCampo)
                    colunas.Add("Movimento", eTipoCampo.Data)
                    colunas.Add("Vencimento", eTipoCampo.Data)
                    colunas.Add("Prorrogacao", eTipoCampo.Data)
                    colunas.Add("Baixa", eTipoCampo.Data)
                    colunas.Add("Debito", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("Credito", eTipoCampo.ValorSemTotalizador)

                    GerarExcelMovtoBancario(ds, "Titulos", colunas)
                Else
                    GerarExcel()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudar_Click(sender As Object, e As EventArgs) Handles lnkAjudar.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeTitulos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkMovtoBancario_CheckedChanged(sender As Object, e As EventArgs) Handles chkMovtoBancario.CheckedChanged
        lnkExcel.Enabled = Not CType(sender, CheckBox).Checked
        lnkPdf.Enabled = Not CType(sender, CheckBox).Checked
    End Sub
End Class
