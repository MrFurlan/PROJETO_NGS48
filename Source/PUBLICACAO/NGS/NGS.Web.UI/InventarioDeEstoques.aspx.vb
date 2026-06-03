Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Xml
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class InventarioDeEstoques
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String
    Dim Porcentagem As String
    Private Empresa() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("InventarioDeEstoques", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaDepositos()
                ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                DdlMes.SelectedValue = Now.Month()
                DdlAno.SelectedValue = Now.Year()
                ddlUnidade.Focus()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/ApuracaoDeCustos.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
              "FROM Clientes C " & vbCrLf &
              "INNER JOIN ClientesXTipos CT " & vbCrLf &
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
              "WHERE CT.Tipo_Id = 050 " & vbCrLf &
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        ddlUnidade.Items.Insert(0, "")
        ddlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              " from Usuarios" & vbCrLf &
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        'Limpar()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf &
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf &
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
              " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            ddlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next

        ddlEmpresa.Items.Insert(0, "")
        ddlEmpresa.SelectedIndex = 0
    End Sub

    Private Sub CargaDepositos()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlDeposito.Items.Clear()

        Sql = " SELECT  Clientes.Cliente_Id AS Codigo, Clientes.Endereco_Id, Clientes.Reduzido, " & vbCrLf &
              "         Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf &
              " FROM    Clientes INNER JOIN" & vbCrLf &
              "         ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id " & vbCrLf &
              "     AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id" & vbCrLf &
              " where   ClientesXTipos.Tipo_Id = 3" & vbCrLf &
              " Order by  Clientes.Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlDeposito.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlDeposito.Items.Insert(0, "")
        DdlDeposito.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        'DdlDeposito.SelectedIndex = 0
        ucSelecaoProduto.Limpar()
        'DdlCusto.SelectedIndex = 0
        'DdlMes.SelectedIndex = 0
        'DdlAno.SelectedIndex = 0
        'txtQuantidade.Text = ""
        'txtValorDaMercadoria.Text = ""
        'txtValorDoFrete.Text = ""
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Function Valida()
        If ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        End If
        If DdlDeposito.Text = "" Then
            MsgBox(Me.Page, "Depósito é obrigatório.")
            Return False
        End If

        Return True
    End Function

    Protected Sub DdlMes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlAno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

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

    Private Function getDataSet() As DataSet
        Dim Ds As New DataSet

        Porcentagem = IIf(txtPorcentagem.Text = "", 100, Replace(txtPorcentagem.Text, ",", "."))

        Sql = "SELECT ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id, " & vbCrLf &
              "       ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf &
              "       ApuracaoDeCustos.Produto_Id, Produtos.Nome AS NomeProduto, " & vbCrLf &
              "       ApuracaoDeCustos.CodigoDeCusto_Id, PlanoDeCustos.Descricao as DescCusto, " & vbCrLf &
              "       ApuracaoDeCustos.Quantidade AS Peso, " & vbCrLf &
              "       (case " & vbCrLf &
              "          when isnull(ApuracaoDeCustos.Quantidade,0) <> 0 and isnull(ApuracaoDeCustos.ValorDoProduto,0) <> 0  " & vbCrLf &
              "            then ISNULL((ApuracaoDeCustos.ValorDoProduto / ApuracaoDeCustos.Quantidade) * Isnull(TabelaDePrecosDeMercado.BaseDeCalculo,1), 0) " & vbCrLf &
              "            else 0.00 " & vbCrLf &
              "       end) AS ValorUnitario, " & vbCrLf &
              "       ApuracaoDeCustos.ValorDoProduto AS ValorDeMercado, " & vbCrLf &
              "       ApuracaoDeCustos.ValorDoFrete, " & vbCrLf &
              "       ISNULL(ApuracaoDeCustos.ValorDoProduto + ApuracaoDeCustos.ValorDoFrete, 0) AS ValorTotal, " & vbCrLf &
              "       Empresa.Nome as EmpresaNome,Empresa.Cidade as EmpresaCidade,Empresa.Estado as EmpresaEstado,Empresa.Reduzido as EmpresaReduzido,  " & vbCrLf &
              "       Deposito.Nome AS DepositoNome,Deposito.Cidade AS DepositoCidade, Deposito.Estado AS DepositoEstado, Deposito.Reduzido AS DepositoReduzido,  " & vbCrLf &
              "       EmpresaDestino_Id as Destino, Destino.Cidade AS DestinoCidade, Destino.Reduzido AS DestinoReduzido, " & vbCrLf &
              "       Isnull(PlanoDeCustos.Desdobramento,'False') as Desdobramento,Isnull(TabelaDePrecosDeMercado.Data_Id,'" & DdlAno.SelectedValue + "/" + DdlMes.SelectedValue & "/01') as Data_Id, " & vbCrLf &
              "       ( " & IIf(RadReal.Checked, "TabelaDePrecosDeMercado.ValorOficial", "TabelaDePrecosDeMercado.ValorMoeda") & "* " & Porcentagem & ")/100 as ValorOficial " & vbCrLf &
              "  Into #Temp " & vbCrLf &
              "  FROM ApuracaoDeCustos  " & vbCrLf &
              "	INNER JOIN Clientes AS Empresa" & vbCrLf &
              "    ON ApuracaoDeCustos.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf &
              "   AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id  " & vbCrLf &
              "	 LEFT JOIN Clientes AS Deposito" & vbCrLf &
              "    ON ApuracaoDeCustos.Deposito_Id    = Deposito.Cliente_Id  " & vbCrLf &
              "   AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id  " & vbCrLf &
              "  LEFT JOIN Clientes AS Destino " & vbCrLf &
              "    ON ApuracaoDeCustos.EmpresaDestino_Id    = Destino.Cliente_Id  " & vbCrLf &
              "   AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id   " & vbCrLf &
              "	INNER JOIN Produtos " & vbCrLf &
              "    ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id " & vbCrLf &
              " INNER JOIN PlanoDeCustos " & vbCrLf &
              "    ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id " & vbCrLf &
              "  Left JOIN TabelaDePrecosDeMercado " & vbCrLf &
              "    ON ApuracaoDeCustos.Empresa_Id     = TabelaDePrecosDeMercado.Empresa_Id  " & vbCrLf &
              "   AND ApuracaoDeCustos.EndEmpresa_Id  = TabelaDePrecosDeMercado.EndEmpresa_Id  " & vbCrLf &
              "   AND ApuracaoDeCustos.Deposito_Id    = TabelaDePrecosDeMercado.Deposito_Id " & vbCrLf &
              "   AND ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id  " & vbCrLf &
              "   AND ApuracaoDeCustos.Produto_Id     = TabelaDePrecosDeMercado.Produto_Id " & vbCrLf &
              "   AND month(Data_Id)                  = " & DdlMes.SelectedValue & vbCrLf &
              "   AND year(Data_Id)                   ='" & DdlAno.SelectedValue & "' " & vbCrLf &
              " WHERE ApuracaoDeCustos.CodigoDeCusto_Id= '920'  " & vbCrLf &
              "   AND ApuracaoDeCustos.Ano_Id = '" & DdlAno.SelectedValue & "'" & vbCrLf &
              "   AND ApuracaoDeCustos.Mes_Id = " & DdlMes.SelectedValue & vbCrLf
        If ddlEmpresa.Text <> "" Then
            campo = ddlEmpresa.SelectedValue.Split("-")
            If ChkTodasAsFiliais.Checked = True Then
                Sql &= "   AND ApuracaoDeCustos.Empresa_ID  Like  '" & campo(0).Substring(0, 8) & "%' " & vbCrLf
            Else
                Sql &= "   And ApuracaoDeCustos.Empresa_ID    = '" & campo(0) & "'" & vbCrLf &
                       "   And ApuracaoDeCustos.EndEmpresa_Id = " & campo(1)
            End If
        End If

        If DdlDeposito.Text <> "" Then
            campo = DdlDeposito.SelectedValue.Split("-")
            Sql &= "   And ApuracaoDeCustos.Deposito_Id    ='" & campo(0) & "'" & vbCrLf &
                   "   And ApuracaoDeCustos.EndDeposito_Id = " & campo(1)
        End If


        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList
            RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "ApuracaoDeCustos.Produto_Id", "")
            Sql &= " AND " & RetornoProdutos(0)
        End If

        Sql &= " Order by ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id, " & vbCrLf &
               "          ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf &
               "          ApuracaoDeCustos.Produto_Id, ApuracaoDeCustos.CodigoDeCusto_Id " & vbCrLf

        Sql &= " Delete #temp " & vbCrLf &
               "  Where Peso = 0 " & vbCrLf &
               "    AND ValorUnitario  = 0 " & vbCrLf &
               "    AND ValorDeMercado = 0 " & vbCrLf &
               "    AND ValorDoFrete   = 0 " & vbCrLf &
               "    AND ValorTotal     = 0 " & vbCrLf

        Sql &= " Select Empresa_Id, NomeProduto,CodigoDeCusto_Id,Destino,Desdobramento, Deposito_Id, EndDeposito_Id, max(data_Id) as Data_Id " & vbCrLf &
               "   Into #Temp1 " & vbCrLf &
               "   From #temp " & vbCrLf &
               "  Group by Empresa_Id, NomeProduto, CodigoDeCusto_Id, Destino, Desdobramento, Deposito_Id, EndDeposito_Id " & vbCrLf

        Sql &= " Select #temp.* " & vbCrLf &
               "   From #Temp " & vbCrLf &
               "  Inner Join #Temp1 " & vbCrLf &
               "     ON #Temp.data_Id          = #temp1.data_id " & vbCrLf &
               "    AND #Temp.Empresa_Id       = #temp1.Empresa_Id " & vbCrLf &
               "    AND #Temp.CodigoDeCusto_Id = #temp1.CodigoDeCusto_Id " & vbCrLf &
               "    AND #Temp.Destino          = #temp1.Destino " & vbCrLf &
               "    AND #Temp.Desdobramento    = #temp1.Desdobramento " & vbCrLf &
               "    AND #Temp.Deposito_Id      = #temp1.Deposito_Id " & vbCrLf &
               "    AND #Temp.EndDeposito_Id   = #temp1.EndDeposito_Id " & vbCrLf &
               "    AND #Temp.NomeProduto      = #temp1.NomeProduto " & vbCrLf
        Ds = Banco.ConsultaDataSet(Sql, "InventarioEstoque")

        Return Ds
    End Function

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorioDados() 'Excel Dados
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "InventarioDeEstoques")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("InventarioDeEstoques", "RELATORIO") Then
                Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
                Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
                Dim ds As DataSet = getDataSet()

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                    Exit Sub
                End If

                Dim param As New Dictionary(Of String, Object)
                param.Add("Empresa", objEmpresa.Nome)
                param.Add("CNPJ", "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                param.Add("INSCRI", "Inscr.Est.: " & objEmpresa.InscricaoEstadual)
                param.Add("Periodo", "Relativo a " & Funcoes.MesPorExtenso(DdlMes.Text) & " de " & DdlAno.SelectedValue)
                param.Add("Consolidado", "Consolidado")
                param.Add("TipoValor", IIf(RadValorReal.Checked, "ValorReal", "ValorAvaliado"))
                param.Add("Emissao", "Emissão : " & Now)

                Funcoes.BindReport(Me.Page, ds, IIf(RadPorProduto.Checked, "Cr_InventarioDeEstoquesPorProduto", "Cr_InventarioDeEstoquesPorFilial"), IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), param)
            Else
                MsgBox(Me.Page, "Usuário sempermissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorioDados()
        Try
            If Funcoes.VerificaPermissao("InventarioDeEstoques", "RELATORIO") Then
                Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
                Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
                Dim data As String = CStr(DdlMes.SelectedValue & "/" & DdlAno.SelectedValue)
                Dim ds As DataSet = getDataSet()

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                    Exit Sub
                End If

                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)

                        'criando planilha 
                        Dim worksheet As ExcelWorksheet

                        'criando título da planilha 
                        If (RadPorFilial.Checked) Then
                            worksheet = package.Workbook.Worksheets.Add("Inventario de Estoques por filial.")
                        Else
                            worksheet = package.Workbook.Worksheets.Add("Inventario de Estoques por produtos.")
                        End If

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
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "INVENTARIO DE ESTOQUES POR DADOS")
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & data)
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        For Each col As DataColumn In ds.Tables(0).Columns
                            If col.ColumnName <> "Row" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            End If
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

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
                            worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                            'formatando células Peso
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000"

                            'formatando células valores
                            worksheet.Cells(String.Format("J{0}:K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

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
                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                'download do arquivo pelo browser
                'Download(fileName)
            Else
                MsgBox(Me.Page, "Usuário sempermissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class