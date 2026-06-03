Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class RelatorioRomaneiosMonsanto
    Inherits BasePage

#Region "Procedimentos"

    Dim Sql As String

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioRomaneiosMonsanto", "ACESSAR") Then
                BuscaUnidadeNegocio()
                BuscarGrupoDeProdutos()
                BuscarOperacoes()
                LiberaEmpresa()

                Limpar()

                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

#Region "Procedimentos"

    Private Sub BuscaUnidadeNegocio()
        ddl.Carregar(cmbUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(cmbUnidadeNegocio, cmbEmpresa)
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidadeNegocio.SelectedValue, True)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbUnidadeNegocio.Enabled = False
            cmbEmpresa.Enabled = False
        End If
    End Sub

    Private Sub BuscarProdutos()
        ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, "Grupo = " & DdlGrupoDeProdutos.SelectedValue, True)
    End Sub

    Private Sub BuscarOperacoes()
        ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "SO.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Private Sub BuscarGrupoDeProdutos()
        Dim ds As New DataSet

        Sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao"
        Sql &= " FROM         GruposDeEstoques INNER JOIN"
        Sql &= "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo"
        Sql &= " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) and Produtos.Agrupar = 'N' order by GruposDeEstoques.Descricao"

        ds = Banco.ConsultaDataSet(Sql, "GruposDeEstoques")

        For Each drGrupo As DataRow In ds.Tables(0).Rows
            DdlGrupoDeProdutos.Items.Add(New ListItem(drGrupo("Descricao"), drGrupo("Grupo_Id")))
        Next

        DdlGrupoDeProdutos.Items.Insert(0, "")
        DdlGrupoDeProdutos.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        txtClientes.Text = ""
        txtCodigoCliente.Value = ""
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
        DdlGrupoDeProdutos.SelectedIndex = 0
        cmbProdutos.Items.Clear()
        chkClientes.Checked = False
        txtPedido.Text = ""
        HID.Value = Guid.NewGuid.ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
    End Sub

    Private Function ValidarCampos() As Boolean

        If cmbEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!", eTitulo.Info)
            Return False

        End If

        If txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida!", eTitulo.Info)
            txtDataInicial.Focus()
            Return False
        End If

        If txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida!", eTitulo.Info)
            txtDataFinal.Focus()
            Return False
        End If

        If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final!", eTitulo.Info)
            txtDataInicial.Focus()
            Return False
        End If

        Return True
    End Function

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objClienteESL" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteESL" & HID.Value), [Lib].Negocio.Cliente))
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteESL" & HID.Value)
        ElseIf Not Session("objPedidosRESL" & HID.Value) Is Nothing Then
            Dim p As Pedido = CType(Session("objPedidosRESL" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = p.Codigo
            Session.Remove("objPedidosRESL" & HID.Value)
        End If
    End Sub


#End Region

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupoDeProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            cmbProdutos.Items.Clear()
            BuscarProdutos()
            cmbProdutos.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscarSubOperacoes(cmbOperacao.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteESL" & HID.Value.ToString(), "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If cmbEmpresa.SelectedIndex = -1 Or txtCodigoCliente.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido")
            Else
                HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
                Dim strCodigo As String() = cmbEmpresa.SelectedValue.Split("-")
                Dim strCodCliente As String() = txtCodigoCliente.Value.Split("-")
                Dim strEmpresa() As String = cmbEmpresa.SelectedValue.ToString.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                Dim parameters As New Dictionary(Of String, Object)
                parameters("unidade") = cmbUnidadeNegocio.SelectedValue
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                Popup.ConsultaDePedidos(Me.Page, "objPedidosRESL" & HID.Value)
                ucConsultaPedidos.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            ImprimirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            ImprimirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioEntradaSaidaLaudo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function RelatorioDataSet() As DataSet
        Sql = " --drop table #Monsanto" & vbCrLf &
              " SELECT RD.Empresa_Id, RD.EndEmpresa_Id, RD.Romaneio_Id, " & vbCrLf &
              "        P.Pesagem_Id AS Laudo, " & vbCrLf &
              "        CASE WHEN ISNULL(N.Nota_Id,0) > 0 THEN N.Nota_Id ELSE 0 END AS NotaFiscal, " & vbCrLf &
              "        R.Movimento, R.Pedido, P.Cliente, P.EndCliente, Cli.Nome as NomeCliente, R.Operacao, R.SubOperacao, SO.Descricao, " & vbCrLf &
              "        Prd.Produto_Id AS Produto, Prd.Nome AS NomeProduto, " & vbCrLf &
              "        CASE WHEN ISNULL(R.Processo,'') = '' THEN 'PESAGEM' ELSE R.Processo END AS Processo, " & vbCrLf &
              "        SUM(R.PesoBruto) AS PesoBruto, " & vbCrLf &
              "        SUM(R.PesoLiquido) AS PesoLiquido, " & vbCrLf &
              "        SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) AS Intacta, " & vbCrLf &
              "        CASE " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 0 THEN 'NÃO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 1 THEN 'NEGATIVO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 2 THEN 'POSITIVO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 3 THEN 'DECLARADO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 4 THEN 'PARTICIPANTE' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 5 THEN 'ORIGEM PARTICIPANTE' " & vbCrLf &
              "        END AS Tipo, " & vbCrLf &
              "        CASE WHEN ISNULL(N.Situacao,0) = 1 THEN 1 ELSE 2 END AS TemNota " & vbCrLf &
              " INTO #Monsanto " & vbCrLf &
              " FROM RomaneiosXDescontos RD " & vbCrLf &
              " INNER JOIN Romaneios R ON R.Empresa_Id = RD.Empresa_Id AND R.EndEmpresa_Id = RD.EndEmpresa_Id AND R.Romaneio_Id = RD.Romaneio_Id " & vbCrLf &
              " INNER JOIN RomaneiosXPesagens RP ON RP.Empresa_Id = R.Empresa_Id AND RP.EndEmpresa_Id = R.EndEmpresa_Id AND RP.Romaneio_Id = R.Romaneio_Id " & vbCrLf &
              " INNER JOIN Pesagem P ON P.Empresa_Id = RP.Empresa_Id AND P.EndEmpresa_Id = RP.EndEmpresa_Id AND P.Pesagem_Id = RP.Pesagem_Id " & vbCrLf &
              " INNER JOIN SubOperacoes SO ON SO.Operacao_Id = R.Operacao AND SO.Suboperacoes_Id = R.SubOperacao " & vbCrLf &
              " INNER JOIN Produtos Prd ON Prd.Produto_Id = R.Produto " & vbCrLf &
              " INNER JOIN Clientes Cli ON Cli.Cliente_Id = P.Cliente AND Cli.Endereco_Id = P.EndCliente " & vbCrLf &
              " LEFT JOIN NotasFiscaisXRomaneios NR ON NR.Empresa_Id = R.Empresa_Id AND NR.EndEmpresa_Id = R.EndEmpresa_Id AND NR.Romaneio_Id = R.Romaneio_Id " & vbCrLf &
              " LEFT JOIN NotasFiscais N ON N.Empresa_Id = NR.Empresa_Id AND N.EndEmpresa_Id = NR.EndEmpresa_Id AND N.Cliente_Id = NR.Cliente_Id " & vbCrLf &
              "      AND N.EndCliente_Id = NR.EndCliente_Id AND N.EntradaSaida_Id = NR.EntradaSaida_Id AND N.Serie_Id = NR.Serie_Id AND N.Nota_Id = NR.Nota_Id " & vbCrLf &
              " WHERE R.Empresa_Id = '" & cmbEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
              " AND R.EndEmpresa_Id = " & cmbEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
              " AND P.Situacao = 1 " & vbCrLf &
              " AND R.EntradaSaida = '" & IIf(optEntrada.Checked, "E", "S") & "'" & vbCrLf &
              " AND RD.Analise_Id IN (12) " & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtClientes.Text) Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= " AND P.Cliente = '" & strCliente(0) & "'" & vbCrLf
            Sql &= " AND P.EndCliente = " & strCliente(1) & vbCrLf
        End If

        If cmbSubOperacao.SelectedIndex > 0 Then
            Sql &= " AND R.Operacao = " & cmbSubOperacao.SelectedValue.Split("-")(0) & vbCrLf
            Sql &= " AND R.SubOperacao = " & cmbSubOperacao.SelectedValue.Split("-")(1) & vbCrLf
        ElseIf cmbOperacao.SelectedIndex > 0 Then
            Sql &= " AND R.Operacao = " & cmbOperacao.SelectedValue & vbCrLf
        End If

        If cmbProdutos.SelectedIndex > 0 Then
            Sql &= " AND Prd.Produto_Id = '" & cmbProdutos.SelectedValue & "'" & vbCrLf
        ElseIf DdlGrupoDeProdutos.SelectedIndex > 0 Then
            Sql &= " AND Prd.Grupo = '" & DdlGrupoDeProdutos.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Sql &= " AND R.Pedido = " & txtPedido.Text & vbCrLf
        Else
            Sql &= " AND R.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
        End If

        Sql &= " GROUP BY RD.Empresa_Id, RD.EndEmpresa_Id, RD.Romaneio_Id, P.Pesagem_Id, N.Nota_Id, R.Movimento, R.Pedido, P.Cliente, P.EndCliente, Cli.Nome, R.Operacao, R.SubOperacao, SO.Descricao, Prd.Produto_Id, Prd.Nome, R.Processo, N.Situacao; " & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              " INSERT INTO #Monsanto (Empresa_Id, EndEmpresa_Id, Romaneio_ID, Laudo, NotaFiscal, Movimento, Pedido, Cliente, EndCliente, NomeCliente, Operacao, SubOperacao, Descricao, Produto, NomeProduto, Processo, PesoBruto, PesoLiquido, Intacta, Tipo, TemNota) " & vbCrLf &
              " SELECT RD.Empresa_Id, RD.EndEmpresa_Id, RD.Romaneio_Id, 0 AS Laudo, N.Nota_Id AS NotaFiscal, " & vbCrLf &
              "        R.Movimento, R.Pedido, N.Cliente_Id as Cliente, N.EndCliente_Id as EndCliente, Cli.Nome as NomeCliente, R.Operacao, R.SubOperacao, SO.Descricao, " & vbCrLf &
              "        Prd.Produto_Id AS Produto, Prd.Nome AS NomeProduto, " & vbCrLf &
              "        CASE WHEN ISNULL(R.Processo,'') = '' THEN 'PESAGEM' ELSE R.Processo END AS Processo, " & vbCrLf &
              "        SUM(R.PesoBruto) AS PesoBruto, SUM(R.PesoLiquido) AS PesoLiquido, " & vbCrLf &
              "        SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) AS Intacta, " & vbCrLf &
              "        CASE " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 0 THEN 'NÃO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 1 THEN 'NEGATIVO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 2 THEN 'POSITIVO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 3 THEN 'DECLARADO' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 4 THEN 'PARTICIPANTE' " & vbCrLf &
              "             WHEN SUM(CASE WHEN RD.Analise_Id = 12 THEN Percentual ELSE 0 END) = 5 THEN 'ORIGEM PARTICIPANTE' " & vbCrLf &
              "        END AS Tipo, 1 AS TemNota " & vbCrLf &
              " FROM RomaneiosXDescontos RD " & vbCrLf &
              " INNER JOIN Romaneios R ON R.Empresa_Id = RD.Empresa_Id AND R.EndEmpresa_Id = RD.EndEmpresa_Id AND R.Romaneio_Id = RD.Romaneio_Id " & vbCrLf &
              " INNER JOIN NotasFiscaisXRomaneios NR ON NR.Empresa_Id = R.Empresa_Id AND NR.EndEmpresa_Id = R.EndEmpresa_Id AND NR.Romaneio_Id = R.Romaneio_Id " & vbCrLf &
              " INNER JOIN NotasFiscais N ON N.Empresa_Id = NR.Empresa_Id AND N.EndEmpresa_Id = NR.EndEmpresa_Id AND N.Cliente_Id = NR.Cliente_Id " & vbCrLf &
              "      AND N.EndCliente_Id = NR.EndCliente_Id AND N.EntradaSaida_Id = NR.EntradaSaida_Id AND N.Serie_Id = NR.Serie_Id AND N.Nota_Id = NR.Nota_Id " & vbCrLf &
              " INNER JOIN SubOperacoes SO ON SO.Operacao_Id = N.Operacao AND SO.Suboperacoes_Id = N.SubOperacao " & vbCrLf &
              " INNER JOIN Produtos Prd ON Prd.Produto_Id = R.Produto " & vbCrLf &
              " INNER JOIN Clientes Cli ON Cli.Cliente_Id = N.Cliente_Id AND Cli.Endereco_Id = N.EndCliente_Id " & vbCrLf &
              " WHERE R.Empresa_Id = '" & cmbEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
              " AND R.EndEmpresa_Id = " & cmbEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
              " AND R.Processo = 'NOTA FISCAL'" & vbCrLf &
              " AND RD.Analise_Id IN (12) " & vbCrLf &
              " AND R.EntradaSaida = '" & IIf(optEntrada.Checked, "E", "S") & "'" & vbCrLf &
              " AND N.Situacao = 1 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtClientes.Text) Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= " AND N.Cliente_Id = '" & strCliente(0) & "'" & vbCrLf
            Sql &= " AND N.EndCliente_Id = " & strCliente(1) & vbCrLf
        End If

        If cmbSubOperacao.SelectedIndex > 0 Then
            Sql &= " AND R.Operacao = " & cmbSubOperacao.SelectedValue.Split("-")(0) & vbCrLf
            Sql &= " AND R.SubOperacao = " & cmbSubOperacao.SelectedValue.Split("-")(1) & vbCrLf
        ElseIf cmbOperacao.SelectedIndex > 0 Then
            Sql &= " AND R.Operacao = " & cmbOperacao.SelectedValue & vbCrLf
        End If

        If cmbProdutos.SelectedIndex > 0 Then
            Sql &= " AND Prd.Produto_Id = '" & cmbProdutos.SelectedValue & "'" & vbCrLf
        ElseIf DdlGrupoDeProdutos.SelectedIndex > 0 Then
            Sql &= " AND Prd.Grupo = '" & DdlGrupoDeProdutos.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Sql &= " AND R.Pedido = " & txtPedido.Text & vbCrLf
        Else
            Sql &= " AND R.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
        End If

        Sql &= " GROUP BY RD.Empresa_Id, RD.EndEmpresa_Id, RD.Romaneio_Id, N.Nota_Id, R.Movimento, R.Pedido, N.Cliente_Id, N.EndCliente_Id, Cli.Nome, R.Operacao, R.SubOperacao, SO.Descricao, Prd.Produto_Id, Prd.Nome, R.Processo; " & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              " SELECT Empresa_Id, EndEmpresa_Id, Romaneio_ID, Laudo, NotaFiscal, Movimento, Pedido, Cliente, EndCliente, NomeCliente, Operacao, SubOperacao, Descricao, Produto, NomeProduto, Processo, PesoBruto, PesoLiquido, Intacta, Tipo, " & vbCrLf &
              "        CASE WHEN TemNota = 1 THEN 'SIM' ELSE 'NÃO' END AS TemNota " & vbCrLf &
              " FROM #Monsanto " & vbCrLf &
              " ORDER BY Movimento;"

        Dim banco As New AcessaBanco
        Return banco.ConsultaDataSet(Sql, "Monsanto")
    End Function

    Private Sub PlanilhaRelatorioRomaneiosMonsanto(ds As DataSet, package As ExcelPackage)
        Dim objEmpresa As Cliente = New Cliente(cmbEmpresa.SelectedValue.Split("-")(0))

        'criando planilha títulos
        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("RomaneiosMonsanto")

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
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RomaneiosMonsanto")
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha que informa o período selecionado na página
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)) & " - " & IIf((optEntrada.Checked), "ENTRADAS", "SAÍDAS"))
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        rowIndex += 1

        'criando linha com o cabeçalho da planilha
        For Each col As DataColumn In ds.Tables(0).Columns
            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
            columnIndex += 1
        Next

        'criando auto filtro na planilha
        worksheet.Cells("A5:R" & rowIndex).AutoFilter = True

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
                If col.ColumnName = "PesoBruto" OrElse col.ColumnName = "PesoLiquido" Then
                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToDecimal(row(col.ColumnName))
                Else
                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                End If
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

        '**************************************************************************************************************************
        'Q -TOTAL
        'criando colunas de totalizadores
        worksheet.Cells(String.Format("M{0}:O{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("M{0}:O{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("M{0}:O{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("M{0}:O{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("M{0}:O{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        worksheet.Cells(String.Format("M{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

        '**************************************************************************************************************************
        'N -PesoBruto
        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("N{0}", rowIndex)).Formula = String.Format("=SUM(N6:N{0})", rowIndex - 1)
        worksheet.Cells(String.Format("N6:N{0}", rowIndex)).Style.Numberformat.Format = "0_ ;[Red]-0"

        '**************************************************************************************************************************
        'O -PesoLiquido
        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Bold = True
        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
        worksheet.Cells(String.Format("O{0}", rowIndex)).Formula = String.Format("=SUM(O6:O{0})", rowIndex - 1)
        worksheet.Cells(String.Format("O6:O{0}", rowIndex)).Style.Numberformat.Format = "0_ ;[Red]-0"

        'setando autofit nas células da planilha
        worksheet.Cells.AutoFitColumns(0)
    End Sub

    Private Shared Function getParametersConsulta(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String,
                                                  ByVal CodigoCliente As String, ByVal EndCliente As Integer,
                                                  ByVal DataLimite As Date?, ByVal Entrada As Boolean, ByVal Saida As Boolean, ByVal MostraDescProd As Boolean) As String
        Dim obj As Cliente = New Cliente(CodigoEmpresa, EndEmpresa)

        Dim param As String = "Empresa: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " / " & obj.Estado.Codigo

        If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
            obj = New Cliente(CodigoCliente, EndCliente)
            param &= vbCrLf & "Cliente: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " / " & obj.Estado.Codigo
        End If

        If MostraDescProd Then
            param &= " - Parâmetros: Mostrar descrição do produto"
        End If

        If DataLimite IsNot Nothing Then
            param &= " - Data Limite: " & CDate(DataLimite).ToString("dd/MM/yyyy")
        End If

        If Not String.IsNullOrWhiteSpace(CodigoPedido) Then
            param &= " - Pedido(s): " & CodigoPedido
        End If

        If Entrada OrElse Saida Then
            If Entrada AndAlso Saida Then
                param &= " - Entrada/Saída: Entrada e Saída"
            Else
                param &= " - Entrada/Saída: " & IIf(Entrada, "Entrada", "Saída")
            End If
        End If

        Return param
    End Function

    Private Sub ImprimirRelatorio(ByVal pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioRomaneiosMonsanto", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim ds As DataSet = RelatorioDataSet()

                    If ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Não existem dados para gerar o relatório!", eTitulo.Info)
                        Exit Sub
                    End If

                    If pdf Then
                        Dim empresa As Cliente = New Cliente(cmbEmpresa.SelectedValue.Split("-")(0))
                        Dim clienteId As String = ""
                        Dim clienteEnd As Integer = 0
                        Dim pedidoId As Integer = 0
                        If cmdConsultaCliente.Text <> ">" Then
                            clienteId = txtClientes.Text.Split("-")(0)
                            clienteEnd = txtClientes.Text.Split("-")(1)
                        End If
                        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                            pedidoId = txtPedido.Text
                        End If
                        Dim param As New Dictionary(Of String, Object)
                        param.Add("ParametersConsulta", getParametersConsulta(empresa.Codigo, empresa.CodigoEndereco, pedidoId, clienteId, clienteEnd,
                                                                              txtDataFinal.Text, optEntrada.Checked, optSaida.Checked, False))

                        Funcoes.BindReport(Page, ds, "Cr_RelatorioRomaneiosMonsanto", eExportType.PDF, param, False, "", empresa.Codigo, empresa.CodigoEndereco)
                    Else
                            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                        If File.Exists(fileName) Then
                            File.Delete(fileName)
                        End If

                        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                            Using package As New ExcelPackage(arquivo)
                                PlanilhaRelatorioRomaneiosMonsanto(ds, package)
                                package.Save()
                            End Using
                        End Using
                        Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

End Class