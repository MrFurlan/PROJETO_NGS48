Imports System.Data
Imports System.Drawing
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CustoMedio
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String
    Dim Empresa() As String
    Dim Deposito() As String
    Dim Porcentagem As String
    Dim Sqla As String
    Dim i As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim scriptManager As ScriptManager = ScriptManager.GetCurrent(Me)
        scriptManager.AsyncPostBackTimeout = 3600 * 5
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CustoMedio", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaDepositos()
                ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                ddlUnidade.Focus()
                Limpar()
                DdlAno.SelectedValue = Format(Today, "yyyy")
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
        ddlEmpresa.Items.Clear()

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
        ucSelecaoProduto.Limpar()

        LiberaEmpresa()
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
        Return True
    End Function

    Protected Sub DdlMes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function GetParameters() As String
        Dim Par As String = String.Empty

        Par = "Parâmetros:" & vbCrLf
        Par &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf

        If DdlDeposito.SelectedIndex > 0 Then
            Par &= "Deposito: " & DdlDeposito.SelectedItem.Text & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Par &= ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "NFI.Produto_id")(1)
        End If

        Par &= "Período de " & Funcoes.MesPorExtenso(DdlMes.Text) & " de " & DdlAno.SelectedValue & vbCrLf

        If chkAuditoria.Checked Then
            Par &= "Auditoria de " & IIf(rdNotas.Checked, rdNotas.Text, IIf(rdRazao.Checked, rdRazao.Text, rdProducao.Text))
        End If

        Return Par

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

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("CustoMedio", "RELATORIO") Then

                If Valida() Then

                    Dim Ds As New DataSet
                    Dim Ds2 As New DataSet

                    Empresa = ddlEmpresa.SelectedValue.Split("-")
                    Deposito = DdlDeposito.SelectedValue.Split("-")

                    Dim Caminho As String = ""

                    If chkAuditoria.Checked Then  'Relatorios Auditoria
                        If rdRazao.Checked Then  'Relatorio Auditoria Razão - Custo
                            Sql = GetSqlAuditoriaRazao()
                        ElseIf rdNotas.Checked Then 'Relatorio Auditoria Notas - Custo
                            Sql = GetSqlAuditoriaPorNota()
                        ElseIf rdProducao.Checked Then 'Relatorio Auditoria Produção - Custo
                            Sql = GetSqlAuditoriaPorProducao()
                        End If

                        Ds = Banco.ConsultaDataSet(Sql, "CustoMedio")

                        If Ds.Tables.Count = 2 Then
                            Dim tableName As String = ""

                            If rdNotas.Checked Then
                                tableName = "CustoMedioAuditoria102"
                            ElseIf rdProducao.Checked Then
                                tableName = "CustoMedioAuditoriaProducao"
                            Else
                                tableName = "CustoMedioAuditoriaRazao"
                            End If

                            Ds.Tables(0).TableName = tableName
                            Ds.Tables(1).TableName = "CustoMedio"
                        End If

                        Caminho = IIf(rdNotas.Checked, "Cr_CustoMedioAuditoria102", IIf(rdRazao.Checked, "Cr_CustoMedioAuditoriaRazao", "Cr_CustoMedioAuditoriaProducao"))
                    Else
                        'Relatorio Padrão
                        Ds = Banco.ConsultaDataSet(GetSqlPadrao, "CustoMedio")
                        Caminho = "Cr_CustoMedio"
                    End If

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Empresa", HttpContext.Current.Session("ssNomeEmpresa"))
                    parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))
                    parameters.Add("Parametros", GetParameters())

                    Funcoes.BindReport(Me.Page, Ds, Caminho, IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorioDados()
        Try
            If Funcoes.VerificaPermissao("CustoMedio", "RELATORIO") Then
                Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
                Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
                Dim data As String = CStr(DdlMes.SelectedItem.Text & "/" & DdlAno.SelectedValue)
                Dim ds As New DataSet
                Dim dt As DataTable = New DataTable()
                Dim titulo As String = ""

                If chkAuditoria.Checked Then  'Relatorios Auditoria
                    If rdRazao.Checked Then  'Relatorio Auditoria Razão - Custo
                        Sql = GetSqlAuditoriaRazao()
                        titulo = "Relatorio Auditoria Razão - Custo"
                    ElseIf rdNotas.Checked Then 'Relatorio Auditoria Notas - Custo
                        Sql = GetSqlAuditoriaPorNota()
                        titulo = "Relatorio Auditoria Notas - Custo"
                    ElseIf rdProducao.Checked Then 'Relatorio Auditoria Produção - Custo
                        Sql = GetSqlAuditoriaPorProducao()
                        titulo = "Relatorio Auditoria Notas - Custo"
                    End If

                    ds = Banco.ConsultaDataSet(Sql, "CustoMedio")
                Else
                    titulo = "Relatório de Custo Medio."
                    ds = Banco.ConsultaDataSet(GetSqlPadrao, "CustoMedio")
                End If

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                    Exit Sub
                End If

                dt = ds.Tables(0)

                If chkAuditoria.Checked Then
                    If rdRazao.Checked Then
                        dt.Columns.Remove("EndEmpresa_Id")
                        dt.Columns.Remove("EmpresaNome")
                        dt.Columns.Remove("EmpresaCidade")
                        dt.Columns.Remove("EmpresaEstado")
                        dt.Columns.Remove("EmpresaReduzido")
                    End If
                Else
                    dt.Columns.Remove("EmpresaNome")
                    dt.Columns.Remove("EmpresaCidade")
                    dt.Columns.Remove("EmpresaEstado")
                    dt.Columns.Remove("EmpresaReduzido")
                    dt.Columns.Remove("DepositoNome")
                    dt.Columns.Remove("DepositoCidade")
                    dt.Columns.Remove("DepositoEstado")
                    dt.Columns.Remove("DepositoReduzido")
                    dt.Columns.Remove("Destino")
                    dt.Columns.Remove("DestinoCidade")
                    dt.Columns.Remove("DestinoReduzido")
                    dt.Columns.Remove("Desdobramento")
                    dt.Columns.Remove("Data_Id")
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
                        worksheet = package.Workbook.Worksheets.Add(titulo)

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
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", titulo)
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

                        For Each col As DataColumn In dt.Columns
                            If col.ColumnName <> "Row" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            End If
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

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

                            If chkAuditoria.Checked Then
                                If rdNotas.Checked Then
                                    'formatando células datas
                                    worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                                    'formatando células Peso
                                    worksheet.Cells(String.Format("W{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000"
                                    'formatando células valores
                                    worksheet.Cells(String.Format("X{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                ElseIf rdRazao.Checked Then
                                    'formatando células datas
                                    worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                                    'formatando células Peso
                                    worksheet.Cells(String.Format("O{0}:P{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000"
                                    'formatando células valores
                                    worksheet.Cells(String.Format("M{0}:N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                ElseIf rdProducao.Checked Then
                                    'formatando células datas   
                                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                                    'formatando células Peso
                                    worksheet.Cells(String.Format("L{0}:M{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000"
                                End If
                            Else
                                'formatando células Peso
                                worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000"
                                'formatando células valores
                                worksheet.Cells(String.Format("J{0}:M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            End If


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
                MsgBox(Me.Page, "Usuário sempermissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function GetSqlAuditoriaRazao() As String
        Dim sql As String = String.Empty
        sql = "SELECT * from ApuracaoDeCustosXFiltroRazao"
        Sqla = "; "
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            i = 0
            If Dr("Processo") = "D" Then
                Sqla &= "Delete #Razao where "
            End If
            If Dr("Processo") = "U" Then
                Sqla &= " Update #Razao " & vbCrLf &
                        " Set Produto = '" & Dr("Produto") & "'" & vbCrLf &
                        " where " & vbCrLf
            End If
            If Dr("Empresa_Id") <> "" Then
                If Dr("EmpresaSinal_Id") = "LIKE" Then
                    Sqla &= " Empresa_Id Like ('" & Dr("Empresa_Id") & "%')"
                    i += 1
                Else
                    Sqla &= " Empresa_Id " & Dr("EmpresaSinal_Id") & " '" & Dr("Empresa_Id") & "'"
                    i += 1
                End If
            End If
            If Dr("Conta_Id") <> "" Then
                If Dr("ContaSinal_Id") = "LIKE" Then
                    Sqla &= IIf(i = 0, " Conta_Id LIKE " & "('" & Dr("Conta_Id") & "%')", " And Conta_Id  LIKE " & "('" & Dr("Conta_Id") & "%')")
                    i += 1
                Else
                    Sqla &= IIf(i = 0, " Conta_Id " & Dr("ContaSinal_Id") & " '" & Dr("Conta_Id") & "'", " And Conta_Id " & Dr("ContaSinal_Id") & " '" & Dr("Conta_Id") & "'")
                    i += 1
                End If
            End If
            If Dr("Produto_Id") <> "" Then
                If Dr("ProdutoSinal_Id") = "LIKE" Then
                    Sqla &= IIf(i = 0, " Produto LIKE " & "('" & Dr("Produto_Id") & "%')", " And Produto  LIKE " & "('" & Dr("Produto_Id") & "%')")
                    i += 1
                Else
                    Sqla &= IIf(i = 0, " Produto " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'", " And Produto " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'")
                    i += 1
                End If
            End If
            If Dr("Lote_Id") > 0 Then
                Sqla &= IIf(i = 0, " Lote_Id " & Dr("LoteSinal_Id") & " " & Dr("Lote_Id"), " And Lote_Id " & Dr("LoteSinal_Id") & " " & Dr("Lote_Id"))
                i += 1
            End If
            If Dr("Placus_Id") > 0 Then
                Sqla &= IIf(i = 0, " Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"), " And Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"))
            End If
            If Dr("Sql") <> "" Then
                Sqla &= IIf(i = 0, " " & Dr("sql"), " And " & Dr("sql"))
            End If
            Sqla &= "; "
        Next

        sql = " SELECT	 " & vbCrLf &
              " Razao.Empresa_Id,  " & vbCrLf &
              " Razao.EndEmpresa_Id,   " & vbCrLf &
              " Empresa.Nome as EmpresaNome,Empresa.Cidade as EmpresaCidade,Empresa.Estado as EmpresaEstado,Empresa.Reduzido as EmpresaReduzido," & vbCrLf &
              " Razao.Conta_Id,  " & vbCrLf &
              " Razao.Lote_Id,  " & vbCrLf &
              " PlanoDeCustosXOrigem.Codigo_Id, " & vbCrLf &
              " Razao.Produto, " & vbCrLf &
              " Razao.Historico, " & vbCrLf &
              " isnull(Razao.Deposito, Razao.Empresa_Id) AS Deposito_Id,  " & vbCrLf &
              " isnull(Razao.EndDeposito, Razao.EndEmpresa_Id)AS EndDeposito_Id,  " & vbCrLf &
              " Deposito.Nome as DepositoNome,Deposito.Cidade as DepositoCidade,Deposito.Estado as DepositoEstado,Deposito.Reduzido as DepositoReduzido," & vbCrLf &
              " Sum(Razao.DebitoOficial) As  " & vbCrLf &
              " DebitoOficial,  " & vbCrLf &
              " Sum(Razao.CreditoOficial) AS  " & vbCrLf &
              " CreditoOficial, " & vbCrLf &
              " Sum(Razao.CreditoQuantidade) As CreditoQuantidade,  " & vbCrLf &
              " Sum(Razao.DebitoQuantidade) AS DebitoQuantidade, " & vbCrLf &
              " Movimento_Id " & vbCrLf &
              " INTO #Razao " & vbCrLf &
              "  FROM   Razao  " & vbCrLf &
              " INNER JOIN PlanoDeCustosXOrigem  " & vbCrLf &
              "     ON Razao.Conta_Id LIKE PlanoDeCustosXOrigem.Conta_Id + '%' " & vbCrLf &
              "  INNER JOIN Produtos  " & vbCrLf &
              "     ON Razao.Produto = Produtos.Produto_Id " & vbCrLf &
              "  INNER JOIN GruposDeEstoques " & vbCrLf &
              "     ON Produtos.Grupo = GruposDeEstoques.Grupo_Id " & vbCrLf &
              "			INNER JOIN  Clientes AS Empresa ON Razao.Empresa_Id = Empresa.Cliente_Id " & vbCrLf &
              "                        AND Razao.EndEmpresa_Id = Empresa.Endereco_Id  " & vbCrLf &
              "			LEFT OUTER JOIN  Clientes AS Deposito ON Razao.Deposito = Deposito.Cliente_Id  " & vbCrLf &
              "                        AND Razao.EndDeposito = Deposito.Endereco_Id  " & vbCrLf &
              "  WHERE (YEAR(Razao.Movimento_Id) = '" & DdlAno.SelectedValue & "')  " & vbCrLf &
              " AND (MONTH(Razao.Movimento_Id) = " & DdlMes.SelectedValue & ")  " & vbCrLf &
              " AND (Empresa_Id like '" & Left(Empresa(0), 8) & "%') " & vbCrLf &
              " And GruposDeEstoques.custo = 1 " & vbCrLf &
              " And Razao.Produto <> ''  " & vbCrLf &
              " And  Razao.Lote_Id not in('7000')  " & vbCrLf &
              " Group By " & vbCrLf &
              " Movimento_Id, " & vbCrLf &
              " Razao.Lote_Id, " & vbCrLf &
              " Razao.Empresa_Id, Razao.EndEmpresa_Id, " & vbCrLf &
              " isnull(Razao.Deposito, Razao.Empresa_Id) ,  " & vbCrLf &
              " isnull(Razao.EndDeposito, Razao.EndEmpresa_Id),  " & vbCrLf &
              " PlanoDeCustosXOrigem.Codigo_Id, Razao.Produto, Razao.Conta_Id, Razao.Historico, " & vbCrLf &
              " Empresa.Nome, Empresa.Cidade, Empresa.Estado, Empresa.Reduzido, " & vbCrLf &
              " Deposito.Nome, Deposito.Cidade, Deposito.Estado, Deposito.Reduzido " & vbCrLf
        sql &= Sqla & "; "
        sql &= " Select * from #Razao where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            sql &= " AND #Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   " AND #Razao.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlDeposito.SelectedValue) Then
            sql &= " AND #Razao.Deposito_Id = '" & DdlDeposito.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   " AND #Razao.EndDeposito_Id = " & DdlDeposito.SelectedValue.Split("-")(1) & vbCrLf
        End If

        'If Not String.IsNullOrWhiteSpace(DdlProduto.SelectedValue) Then
        '    Sql &= " AND #Razao.Produto = '" & DdlProduto.SelectedValue & "'" & vbCrLf
        'End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("GruposDeEstoques.Grupo_Id", "#Razao.Produto")(0)
        End If

        sql &= GetSqlResumo()


        sql &= "Select #temp.* " & vbCrLf &
                           "  From #Temp " & vbCrLf &
                           " INNER JOIN #Razao                                                                                                         " & vbCrLf &
                           " 	ON #Razao.Empresa_Id = #Temp.Empresa_Id                                                                                 " & vbCrLf &
                           " 	AND #Razao.EndEmpresa_Id = #Temp.EndEmpresa_Id                                                                          " & vbCrLf &
                           " 	AND #Razao.Codigo_Id = #Temp.CodigoDeCusto_Id                                                                           " & vbCrLf &
                           " 	AND #Razao.Produto = #Temp.Produto_Id                                                                                   " & vbCrLf &
                           " 	AND #Razao.Deposito_Id = #Temp.Deposito_Id                                                                              " & vbCrLf &
                           " 	AND #Razao.EndDeposito_Id = #Temp.EndDeposito_Id                                                                        " & vbCrLf &
                           "                                                                                                                           " & vbCrLf &
                           " group by #Temp.Empresa_Id, #Temp.endEmpresa_Id, #Temp.Produto_Id,                                                         " & vbCrLf &
                           " 			#Temp.DescCusto, #Temp.Peso, #Temp.ValorUnitario, #Temp.NomeProduto, #Temp.CodigoDeCusto_Id,                    " & vbCrLf &
                           " 			#Temp.Destino, #Temp.Desdobramento, #Temp.Deposito_Id, #Temp.EndDeposito_Id, #Temp.ValorDeMercado,              " & vbCrLf &
                           " 			#Temp.ValorDoFrete, #Temp.ValorTotal, #Temp.EmpresaNome, #Temp.EmpresaCidade, #Temp.EmpresaEstado,              " & vbCrLf &
                           " 			#Temp.EmpresaReduzido, #Temp.DepositoNome, #Temp.DepositoCidade, #Temp.DepositoEstado, #Temp.DepositoReduzido,  " & vbCrLf &
                           " 			#Temp.DestinoCidade, #Temp.DestinoReduzido, #Temp.Data_Id                                                       " & vbCrLf &
                           "" & vbCrLf &
                           " order by #temp.CodigoDeCusto_Id	" & vbCrLf



        Return sql
    End Function

    Private Function GetSqlAuditoriaPorNota() As String
        Dim sql As String = String.Empty

        Dim sqlWhere As String

        i = 0

        sql = "SELECT * from ApuracaoDeCustosXFiltroNotas order by ordem"
        Sqla = "; "
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            sqlWhere = ""
            i = 0
            If Dr("Processo") = "D" Then
                Sqla &= "Delete #Notas where "
            End If
            If Dr("Processo") = "U" Then
                Sqla &= " Update #Notas Set "
                If Dr("Produto") <> "" Then
                    Sqla &= " Produto_Id = '" & Dr("Produto") & "'"
                End If
                If Dr("Empresa_Id") <> Dr("Empresa") And Dr("Empresa").ToString().Length = 14 Then
                    If Dr("Produto") <> "" Then
                        Sqla &= ", "
                    End If
                    Sqla &= " Empresa_Id     ='" & Dr("Empresa") & "'," & vbCrLf &
                            " EndEmpresa_Id  = " & Dr("EndEmpresa") & "," & vbCrLf

                    If Dr("Deposito_Id") <> Dr("Deposito") And Dr("Empresa").ToString().Length = 14 Then
                        Sqla &= " Deposito_Id    ='" & Dr("Deposito") & "',"
                        Sqla &= " EndDeposito_Id = " & Dr("EndDeposito")
                    End If
                    If Dr("Deposito_Id") <> "" Then
                        sqlWhere &= " and Deposito_id = '" & Dr("Deposito_Id") & "'"
                        sqlWhere &= " and EndDeposito_id = '" & Dr("EndDeposito_id") & "'"
                    End If
                ElseIf Dr("Deposito_Id") <> Dr("Deposito") And Dr("Empresa").ToString().Length = 14 Then
                    If Dr("Produto") <> "" Then
                        Sqla &= ", "
                    End If
                    Sqla &= " Deposito_Id    = '" & Dr("Deposito") & "'," & vbCrLf &
                            " EndDeposito_Id = " & Dr("EndDeposito") & vbCrLf

                    If Dr("Deposito_Id") <> "" Then
                        sqlWhere &= " and Deposito_id = '" & Dr("Deposito_Id") & "'" & vbCrLf &
                                    " and EndDeposito_id = " & Dr("EndDeposito_id") & vbCrLf
                    End If
                End If
                Sqla &= " where "
            End If
            If Dr("Empresa_Id") <> "" Then
                If Dr("EmpresaSinal_Id") = "LIKE" Then
                    Sqla &= " Empresa_Id Like ('" & Dr("Empresa_Id") & "%')"
                    i += 1
                Else
                    Sqla &= " Empresa_Id " & Dr("EmpresaSinal_Id") & " '" & Dr("Empresa_Id") & "'"
                    i += 1
                End If
            End If
            If Dr("Produto_Id") <> "" Then
                If Dr("ProdutoSinal_Id") = "LIKE" Then
                    Sqla &= IIf(i = 0, " Produto_Id LIKE " & "('" & Dr("Produto_Id") & "%')", " And Produto_Id  LIKE " & "('" & Dr("Produto_Id") & "%')")
                    i += 1
                Else
                    Sqla &= IIf(i = 0, " Produto_Id " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'", " And Produto_Id " & Dr("ProdutoSinal_Id") & " '" & Dr("Produto_Id") & "'")
                    i += 1
                End If
            End If

            If Dr("Placus_Id") > 0 Then
                Sqla &= IIf(i = 0, " Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"), " And Codigo_Id " & Dr("PlacusSinal_Id") & " " & Dr("Placus_Id"))
            End If
            Sqla &= sqlWhere & "; "
        Next

        sql = "Select NotasFiscais.Empresa_Id, " & vbCrLf &
              "       NotasFiscais.EndEmpresa_Id, " & vbCrLf &
              "       Empresa.Nome as EmpresaNome,Empresa.Cidade as EmpresaCidade,Empresa.Estado as EmpresaEstado,Empresa.Reduzido as EmpresaReduzido," & vbCrLf &
              "       NotasFiscais.Cliente_Id, " & vbCrLf &
              "       NotasFiscais.EndCliente_Id, " & vbCrLf &
              "       Cliente.Nome as ClienteNome,Cliente.Cidade as ClienteCidade,Cliente.Estado as ClienteEstado,Cliente.Reduzido as ClienteReduzido," & vbCrLf &
              "       NotasFiscais.EntradaSaida_Id, " & vbCrLf &
              "       NotasFiscais.Serie_Id, " & vbCrLf &
              "       NotasFiscais.Nota_Id, " & vbCrLf &
              "       NotasFiscais.Pedido, " & vbCrLf &
              "       NotasFiscais.DataDaNota, " & vbCrLf &
              "       SubOperacoes.Operacao_Id, " & vbCrLf &
              "       SubOperacoes.SubOperacoes_Id, " & vbCrLf &
              "       NotasFiscais.Movimento, " & vbCrLf &
              "       NotasFiscaisXItens.Produto_Id, " & vbCrLf &
              "       Produtos.Nome AS NomeProduto,  " & vbCrLf &
              "       NotasFiscaisXItens.PesoFiscal AS Quantidade, " & vbCrLf &
              "	      sum(isnull(case " & vbCrLf &
              "		  		       when ((NotasFiscais.EntradaSaida_Id  = 'E') OR (NotasFiscais.EntradaSaida_Id = 'S' And SubOperacoes.Devolucao = 'S')) " & vbCrLf &
              "					     then NotasFiscaisXItens.Valor - (NFxE.COFINS + NFxE.PIS + NFxE.ICMS)" & vbCrLf &
              "					     else NotasFiscaisXItens.Valor   " & vbCrLf &
              "				     end,0)) as ValorDoProduto, " & vbCrLf &
              "	      NotasFiscais.Deposito AS Deposito_Id, " & vbCrLf &
              "	      NotasFiscais.EndDeposito AS EndDeposito_Id," & vbCrLf &
              "	      Deposito.Nome as DepositoNome,Deposito.Cidade as DepositoCidade,Deposito.Estado as DepositoEstado,Deposito.Reduzido as DepositoReduzido," & vbCrLf &
              "	      SubOperacoes.ApuracaoDeCustos AS CodigoDeCusto_Id," & vbCrLf &
              "	      PlanoDeCustos.Descricao as DescCusto, " & vbCrLf &
              "	      Isnull(PlanoDeCustos.Desdobramento,'False') as Desdobramento,  " & vbCrLf &
              "	      NotasFiscaisXItens.Cfop_Id AS CFOP " & vbCrLf &
              "  into #Notas " & vbCrLf &
              "  FROM NotasFiscais  " & vbCrLf &
              " INNER JOIN SubOperacoes  " & vbCrLf &
              "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id  " & vbCrLf &
              "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id  " & vbCrLf &
              " INNER JOIN NotasFiscaisXItens  " & vbCrLf &
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id  " & vbCrLf &
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf &
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf &
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf &
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
              "  LEFT JOIN Produtos  " & vbCrLf &
              "    ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id  " & vbCrLf &
              " INNER JOIN GruposDeEstoques " & vbCrLf &
              "    ON  Produtos.Grupo = GruposDeEstoques.Grupo_Id " & vbCrLf &
              "  LEFT JOIN PlanoDeCustos " & vbCrLf &
              "    ON SubOperacoes.ApuracaoDeCustos = PlanoDeCustos.Codigo_Id  " & vbCrLf &
              "  LEFT JOIN Clientes AS Empresa " & vbCrLf &
              "    ON NotasFiscais.Empresa_Id    = Empresa.Cliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EndEmpresa_Id = Empresa.Endereco_Id  " & vbCrLf &
              "  LEFT JOIN Clientes AS Cliente" & vbCrLf &
              "    ON NotasFiscais.Cliente_Id    = Cliente.Cliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EndCliente_Id = Cliente.Endereco_Id  " & vbCrLf &
              "  LEFT JOIN Clientes AS Deposito " & vbCrLf &
              "    ON NotasFiscais.Deposito    = Deposito.Cliente_Id  " & vbCrLf &
              "   AND NotasFiscais.EndDeposito = Deposito.Endereco_Id " & vbCrLf &
              " INNER JOIN(Select Empresa_Id," & vbCrLf &
              "                   EndEmpresa_Id," & vbCrLf &
              "                   Cliente_Id," & vbCrLf &
              "                   EndCliente_Id," & vbCrLf &
              "                   EntradaSaida_Id," & vbCrLf &
              "                   Serie_Id," & vbCrLf &
              "                   Nota_Id," & vbCrLf &
              "                   Sequencia_Id," & vbCrLf &
              "                   Produto_Id," & vbCrLf &
              "                   sum(case When Encargo_Id = 'PRODUTO'         then Valor else 0 end) as PRODUTO," & vbCrLf &
              "    			      sum(case When Encargo_Id = 'IPI'             then Valor else 0 end) as IPI, " & vbCrLf &
              "                   sum(case When Encargo_Id = 'FRETES'          then Valor else 0 end) as FRETES," & vbCrLf &
              "			  	      sum(case When Encargo_Id = 'DESP.ADUANEIRAS' then Valor else 0 end) as DESPADUANEIRAS," & vbCrLf &
              "				      sum(case When Encargo_Id = 'DESCONTOS'       then Valor else 0 end) as DESCONTOS," & vbCrLf &
              "				      sum(case When Encargo_Id = 'PIS'             then Valor else 0 end) as PIS," & vbCrLf &
              "				      sum(case When Encargo_Id = 'COFINS'          then Valor else 0 end) as COFINS," & vbCrLf &
              "                   sum(case When Encargo_Id like '%ICMS%'       then Valor else 0 end) as ICMS " & vbCrLf &
              "              from NotasFiscaisXEncargos               " & vbCrLf &
              "             Where Encargo_Id IN ('IPI', 'PRODUTO','FRETES','DESP.ADUANEIRAS','DESCONTOS','PIS','COFINS') " & vbCrLf &
              "                or Encargo_Id like '%ICMS%' " & vbCrLf &
              "             Group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Sequencia_Id, Produto_Id  " & vbCrLf &
              "           ) NFxE" & vbCrLf &
              "     ON NotasFiscaisXItens.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
              "  WHERE (YEAR(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento))  = " & DdlAno.SelectedValue & ")" & vbCrLf &
              "    AND (MONTH(isnull(NotasFiscais.DataParaCusto,NotasFiscais.Movimento)) = " & DdlMes.SelectedValue & ")  " & vbCrLf &
              "    AND (NotasFiscais.Empresa_Id like '" & Left(Empresa(0), 8) & "%') " & vbCrLf &
              "    AND (NotasFiscais.Situacao         = 1) " & vbCrLf &
              "    AND (SubOperacoes.ApuracaoDeCustos > 0) " & vbCrLf &
              "    And (GruposDeEstoques.custo        = 1) " & vbCrLf



        'If DdlProduto.Text <> "" Then
        '    Sql &= " And NotasFiscaisXItens.Produto_Id = '" & DdlProduto.SelectedValue & "' "
        'End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("GruposDeEstoques.Grupo_Id", "NotasFiscaisXItens.Produto_Id")(0)
        End If

        sql &= "  Group by  NotasFiscais.Empresa_Id, NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id,  " & vbCrLf &
               "			Empresa.Nome, Empresa.Cidade, Empresa.Estado, Empresa.Reduzido, " & vbCrLf &
               "			NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, Cliente.Nome, Cliente.Cidade, Cliente.Estado, " & vbCrLf &
               "			Cliente.Reduzido, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, " & vbCrLf &
               "			NotasFiscais.Pedido, NotasFiscais.DataDaNota, SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, " & vbCrLf &
               "			NotasFiscais.Movimento, NotasFiscaisXItens.Produto_Id, Produtos.Nome, NotasFiscaisXItens.PesoFiscal, " & vbCrLf &
               "			NotasFiscais.Deposito, NotasFiscais.EndDeposito, Deposito.Nome, Deposito.Cidade, " & vbCrLf &
               "			Deposito.Estado, Deposito.Reduzido, SubOperacoes.ApuracaoDeCustos, PlanoDeCustos.Descricao, " & vbCrLf &
               "			Isnull(PlanoDeCustos.Desdobramento,'False'), NotasFiscaisXItens.Cfop_Id  " & vbCrLf
        sql &= Sqla & "; "
        sql &= " Select * from #Notas Where CodigoDeCusto_Id <> 1 order by DataDaNota, CodigoDeCusto_Id"

        'End If

        sql &= GetSqlResumo()

        sql &= " Select distinct #temp.*                                            " & vbCrLf &
                          "   From #Temp                                              " & vbCrLf &
                          "   Inner Join #Notas                                       " & vbCrLf &
                          " 	ON #Notas.Empresa_Id = #temp.Empresa_Id                " & vbCrLf &
                          " 	AND #Notas.endEmpresa_Id = #temp.endEmpresa_Id         " & vbCrLf &
                          " 	AND #Notas.Deposito_Id = #temp.Deposito_Id             " & vbCrLf &
                          " 	AND #Notas.EndDeposito_Id = #temp.EndDeposito_Id       " & vbCrLf &
                          " 	AND #Notas.Produto_Id = #temp.Produto_Id               " & vbCrLf &
                          " 	AND #Notas.CodigoDeCusto_Id = #temp.CodigoDeCusto_Id   " & vbCrLf

        Return sql

    End Function

    Private Function GetSqlAuditoriaPorProducao() As String
        Dim sql As String = String.Empty

        i = 0

        sql = "SELECT * from ApuracaoDeCustosXFiltroEstoques" & vbCrLf &
              "; " & vbCrLf
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            i = 0
            Sqla &= "Delete #Estoques where "

            If Dr("Empresa_Id") <> "" Then
                Sqla &= " Empresa_Id like '" & Dr("Empresa_Id") & "%'"
                i += 1
            End If
            If Dr("Deposito_Id") <> "" Then
                Sqla &= IIf(i = 0, " Deposito_Id = '" & Dr("Deposito_Id"), " And Deposito_Id = " & Dr("Deposito_Id"))
                i += 1
            End If
            If Dr("Produto_Id") <> "" Then
                Sqla &= IIf(i = 0, " Produto = " & Dr("Produto_Id"), " And Produto = " & Dr("Produto_Id"))
                i += 1
            End If
            If Dr("ProdutoDerivado_Id") <> "" Then
                Sqla &= IIf(i = 0, " ProdutoDerivado_Id = " & Dr("ProdutoDerivado_Id"), " And ProdutoDerivado_Id = " & Dr("ProdutoDerivado_Id"))
                i += 1
            End If
            If Dr("Operacao_Id") > 0 Then
                Sqla &= IIf(i = 0, " Operacao_Id = " & Dr("Operacao_Id"), " And Operacao_Id = " & Dr("Operacao_Id"))
                i += 1
            End If
            If Dr("SubOperacao_Id") > 0 Then
                Sqla &= IIf(i = 0, " SubOperacao_Id = " & Dr("SubOperacao_Id"), " And SubOperacao_Id = " & Dr("SubOperacao_Id"))
            End If
            Sqla &= "; "
        Next

        sql = " SELECT Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, " & vbCrLf &
              "        Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Operacao_Id, " & vbCrLf &
              "        cast(Producao.SubOperacao_Id as varchar) + ' - ' + SubOperacoes.Descricao as SubOperacao_Id, Producao.ProdutoDerivado_Id, SubOperacoes.ApuracaoDeCustos AS Placus_Id,  " & vbCrLf &
              "        isnull(SubOperacoes.ApuracaoDeCustosContraPartida,0) as PlacusContraPartida, Producao.Movimento_Id, " & vbCrLf &
              "        Producao.Entradas AS QuantidadeEntrada," & vbCrLf &
              "        Producao.Saidas AS QuantidadeSaida" & vbCrLf &
              "   Into #Estoques" & vbCrLf &
              "   FROM Producao" & vbCrLf &
              "  INNER JOIN SubOperacoes" & vbCrLf &
              "     ON Producao.Operacao_Id    = SubOperacoes.Operacao_Id " & vbCrLf &
              "    AND Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              "  INNER JOIN Produtos " & vbCrLf &
              "     ON Producao.Produto_Id = Produtos.Produto_Id" & vbCrLf &
              "  INNER JOIN GruposDeEstoques" & vbCrLf &
              "    ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf &
              " WHERE Producao.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'" & vbCrLf &
              "   AND MONTH(Producao.Movimento_Id) =  " & DdlMes.SelectedValue & vbCrLf &
              "   AND YEAR(Producao.Movimento_Id)  =  " & DdlAno.SelectedValue & vbCrLf &
              "   AND Producao.FisicoFiscal_Id     = 2" & vbCrLf &
              "   AND GruposDeEstoques.custo       = 1" & vbCrLf

        'If DdlProduto.Text <> "" Then
        '    Sql &= " And Producao.Produto_Id = '" & DdlProduto.SelectedValue & "' "
        'End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("GruposDeEstoques.Grupo_Id", "Producao.Produto_Id")(0)
        End If


        sql &= Sqla & "; " 'filtros

        sql &= " Select * from #Estoques where 1=1" & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            sql &= " AND #Estoques.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   " AND #Estoques.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlDeposito.SelectedValue) Then
            sql &= " AND #Estoques.Deposito_Id = '" & DdlDeposito.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   " AND #Estoques.EndDeposito_Id = " & DdlDeposito.SelectedValue.Split("-")(1) & vbCrLf
        End If


        'If Not String.IsNullOrWhiteSpace(DdlProduto.SelectedValue) Then
        '    Sql &= " AND #Estoques.Produto_Id = '" & DdlProduto.SelectedValue & "'" & vbCrLf
        'End If


        If ucSelecaoProduto.TemSelecionado Then
            sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("GruposDeEstoques.Grupo_Id", "#Estoques.Produto_Id")(0)
        End If

        sql &= "order by Placus_Id" & vbCrLf

        sql &= GetSqlResumo()

        sql &= "Select distinct #temp.* " & vbCrLf &
                       "  From #Temp " & vbCrLf &
                       "   Inner Join #Estoques                                      " & vbCrLf &
                       "       ON #Estoques.Empresa_Id = #Temp.Empresa_Id            " & vbCrLf &
                       "       And #Estoques.endEmpresa_Id = #Temp.endEmpresa_Id     " & vbCrLf &
                       "       And #Estoques.Deposito_Id = #Temp.deposito_Id         " & vbCrLf &
                       "       And #Estoques.EndDeposito_Id = #Temp.EndDeposito_Id   " & vbCrLf &
                       "       And #Estoques.Produto_Id = #Temp.Produto_Id           " & vbCrLf &
                       "       And #Estoques.Placus_Id = #Temp.CodigoDeCusto_Id      " & vbCrLf


        Return sql
    End Function

    Private Function GetSqlPadrao() As String
        Dim sql As String = String.Empty
        sql = "SELECT   " & vbCrLf &
              "ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id, " & vbCrLf &
              "ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf &
              "ApuracaoDeCustos.Produto_Id, Produtos.Nome AS NomeProduto, " & vbCrLf &
              "ApuracaoDeCustos.CodigoDeCusto_Id, case when ProdutoDerivado_Id = '' then PlanoDeCustos.Descricao " & vbCrLf &
              "else PlanoDeCustos.Descricao +'-'+ ProdutoDerivado_Id +'-'+ Derivado.Nome end as DescCusto, " & vbCrLf &
              "ApuracaoDeCustos.Quantidade AS Peso, " & vbCrLf &
              "(case " & vbCrLf &
              "when isnull(ApuracaoDeCustos.Quantidade,0) <> 0 and (isnull(ApuracaoDeCustos.ValorDoProduto,0) <> 0 Or isnull(ApuracaoDeCustos.ValorDoProduto,0) <> 0) then " & vbCrLf &
              "ISNULL(((isnull(ApuracaoDeCustos.ValorDoProduto, 0) + isnull(ApuracaoDeCustos.ValorDoFrete, 0)) / ApuracaoDeCustos.Quantidade), 0) " & vbCrLf &
              "else " & vbCrLf &
              "0.00 " & vbCrLf &
              "end) AS ValorUnitario, " & vbCrLf &
              "ApuracaoDeCustos.ValorDoProduto AS ValorDeMercado, " & vbCrLf &
              "ApuracaoDeCustos.ValorDoFrete, " & vbCrLf &
              "ISNULL(ApuracaoDeCustos.ValorDoProduto + ApuracaoDeCustos.ValorDoFrete, 0) AS ValorTotal, " & vbCrLf &
              "Empresa.Nome as EmpresaNome,Empresa.Cidade as EmpresaCidade,Empresa.Estado as EmpresaEstado,Empresa.Reduzido as EmpresaReduzido,  " & vbCrLf &
              "Deposito.Nome AS DepositoNome,Deposito.Cidade AS DepositoCidade, Deposito.Estado AS DepositoEstado, Deposito.Reduzido AS DepositoReduzido,  " & vbCrLf &
              "EmpresaDestino_Id as Destino, Destino.Cidade AS DestinoCidade, Destino.Reduzido AS DestinoReduzido, " & vbCrLf &
              "Isnull(PlanoDeCustos.Desdobramento,'False') as Desdobramento, Isnull(TabelaDePrecosDeMercado.Data_Id,'" & DdlAno.SelectedValue + "/" + DdlMes.SelectedValue & "/01') as Data_Id " & vbCrLf &
              " Into #Temp " & vbCrLf &
              "FROM         ApuracaoDeCustos  " & vbCrLf &
              "			INNER JOIN  Clientes AS Empresa ON ApuracaoDeCustos.Empresa_Id = Empresa.Cliente_Id " & vbCrLf &
              "                        AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id  " & vbCrLf &
              "			LEFT OUTER JOIN  Clientes AS Deposito ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id  " & vbCrLf &
              "                        AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id  " & vbCrLf &
              "         LEFT OUTER JOIN  Clientes AS Destino ON ApuracaoDeCustos.EmpresaDestino_Id = Destino.Cliente_Id  " & vbCrLf &
              "                        AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id   " & vbCrLf &
              "			INNER JOIN  Produtos ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id " & vbCrLf &
              "         Left JOIN  Produtos as Derivado ON ApuracaoDeCustos.ProdutoDerivado_Id = Derivado.Produto_Id " & vbCrLf &
              "         INNER JOIN  PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id " & vbCrLf &
              "         Left JOIN   TabelaDePrecosDeMercado ON ApuracaoDeCustos.Empresa_Id = TabelaDePrecosDeMercado.Empresa_Id " & vbCrLf &
              "                        AND  ApuracaoDeCustos.EndEmpresa_Id = TabelaDePrecosDeMercado.EndEmpresa_Id  " & vbCrLf &
              "                        AND  ApuracaoDeCustos.Deposito_Id = TabelaDePrecosDeMercado.Deposito_Id " & vbCrLf &
              "                        AND  ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id  " & vbCrLf &
              "                        AND  ApuracaoDeCustos.Produto_Id = TabelaDePrecosDeMercado.Produto_Id " & vbCrLf &
              "                        AND  month(Data_Id)  = " & DdlMes.SelectedValue & " AND year(Data_Id)= '" & DdlAno.SelectedValue & "' " & vbCrLf &
              " WHERE     ApuracaoDeCustos.Ano_Id = '" & DdlAno.SelectedValue & "' AND ApuracaoDeCustos.Mes_Id = " & DdlMes.SelectedValue & "  " & vbCrLf

        If ddlEmpresa.Text <> "" Then
            sql &= " And ApuracaoDeCustos.Empresa_ID = '" & Empresa(0) & "' " & vbCrLf &
                   " And ApuracaoDeCustos.EndEmpresa_Id = " & Empresa(1) & vbCrLf
        End If
        If DdlDeposito.Text <> "" Then
            sql &= " And ApuracaoDeCustos.Deposito_Id = '" & Deposito(0) & "' " & vbCrLf &
                   " And ApuracaoDeCustos.EndDeposito_Id = " & Deposito(1) & vbCrLf
        End If
        'If DdlProduto.Text <> "" Then
        '    Sql &= " And ApuracaoDeCustos.Produto_Id = '" & DdlProduto.SelectedValue & "' "
        'End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo_Id", "Produtos.Produto_Id")(0)
        End If

        sql &= "" & vbCrLf &
               " Delete #temp " & vbCrLf &
               "  Where Peso           = 0 " & vbCrLf &
               "    AND ValorUnitario  = 0 " & vbCrLf &
               "    AND ValorDeMercado = 0 " & vbCrLf &
               "    AND ValorDoFrete   = 0 " & vbCrLf &
               "    AND ValorTotal     = 0 " & vbCrLf &
               "    AND CodigoDeCusto_Id <> 920 " & vbCrLf &
                "" & vbCrLf &
               " Select Empresa_Id, NomeProduto,CodigoDeCusto_Id,Destino,Desdobramento, Deposito_Id, EndDeposito_Id, max(data_Id) as Data_Id " & vbCrLf &
               "   Into #Temp1  " & vbCrLf &
               "   From #temp  " & vbCrLf &
               "  Group by Empresa_Id, NomeProduto, CodigoDeCusto_Id, Destino, Desdobramento, Deposito_Id, EndDeposito_Id " & vbCrLf &
                "" & vbCrLf &
               "Select #temp.* " & vbCrLf &
               "  From #Temp " & vbCrLf &
               " Inner Join #Temp1 " & vbCrLf &
               "    ON #Temp.data_Id          = #temp1.data_id " & vbCrLf &
               "   AND #Temp.Empresa_Id       =  #temp1.Empresa_Id " & vbCrLf &
               "   AND #Temp.CodigoDeCusto_Id =  #temp1.CodigoDeCusto_Id " & vbCrLf &
               "   AND #Temp.Destino          =  #temp1.Destino " & vbCrLf &
               "   AND #Temp.Desdobramento    =  #temp1.Desdobramento " & vbCrLf &
               "   AND #Temp.Deposito_Id      =  #temp1.Deposito_Id " & vbCrLf &
               "   AND #Temp.EndDeposito_Id   =  #temp1.EndDeposito_Id " & vbCrLf &
               "   AND #Temp.NomeProduto      =  #temp1.NomeProduto " & vbCrLf &
               " Order by Empresa_Id,  EndEmpresa_Id, " & vbCrLf &
               "          Deposito_Id, EndDeposito_Id," & vbCrLf &
               "          Produto_Id,  CodigoDeCusto_Id " & vbCrLf

        Return sql
    End Function

    Private Function GetSqlResumo() As String
        Dim sql As String = String.Empty

        sql &= "" & vbCrLf &
                          "--__________________________________________________________________________________" & vbCrLf &
                          "    Select" & vbCrLf &
                          "ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id, " & vbCrLf &
                          "ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf &
                          "ApuracaoDeCustos.Produto_Id, Produtos.Nome AS NomeProduto, " & vbCrLf &
                          "ApuracaoDeCustos.CodigoDeCusto_Id, case when ProdutoDerivado_Id = '' then PlanoDeCustos.Descricao " & vbCrLf &
                          "else PlanoDeCustos.Descricao +'-'+ ProdutoDerivado_Id +'-'+ Derivado.Nome end as DescCusto, " & vbCrLf &
                          "ApuracaoDeCustos.Quantidade AS Peso, " & vbCrLf &
                          "(case " & vbCrLf &
                           "when isnull(ApuracaoDeCustos.Quantidade,0) <> 0 and (isnull(ApuracaoDeCustos.ValorDoProduto,0) <> 0 Or isnull(ApuracaoDeCustos.ValorDoProduto,0) <> 0) then " & vbCrLf &
                           "ISNULL(((isnull(ApuracaoDeCustos.ValorDoProduto, 0) + isnull(ApuracaoDeCustos.ValorDoFrete, 0)) / ApuracaoDeCustos.Quantidade), 0) " & vbCrLf &
                          "else " & vbCrLf &
                          "0.00 " & vbCrLf &
                          "end) AS ValorUnitario, " & vbCrLf &
                          "ApuracaoDeCustos.ValorDoProduto AS ValorDeMercado, " & vbCrLf &
                          "ApuracaoDeCustos.ValorDoFrete, " & vbCrLf &
                          "ISNULL(ApuracaoDeCustos.ValorDoProduto + ApuracaoDeCustos.ValorDoFrete, 0) AS ValorTotal, " & vbCrLf &
                          "Empresa.Nome as EmpresaNome,Empresa.Cidade as EmpresaCidade,Empresa.Estado as EmpresaEstado,Empresa.Reduzido as EmpresaReduzido,  " & vbCrLf &
                          "Deposito.Nome AS DepositoNome,Deposito.Cidade AS DepositoCidade, Deposito.Estado AS DepositoEstado, Deposito.Reduzido AS DepositoReduzido,  " & vbCrLf &
                          "EmpresaDestino_Id as Destino, Destino.Cidade AS DestinoCidade, Destino.Reduzido AS DestinoReduzido, " & vbCrLf &
                          "Isnull(PlanoDeCustos.Desdobramento,'False') as Desdobramento, Isnull(TabelaDePrecosDeMercado.Data_Id,'" & DdlAno.SelectedValue + "/" + DdlMes.SelectedValue & "/01') as Data_Id " & vbCrLf &
                          " Into #Temp " & vbCrLf &
                          "FROM         ApuracaoDeCustos  " & vbCrLf &
                          "			INNER JOIN  Clientes AS Empresa ON ApuracaoDeCustos.Empresa_Id = Empresa.Cliente_Id " & vbCrLf &
                          "                        AND ApuracaoDeCustos.EndEmpresa_Id = Empresa.Endereco_Id  " & vbCrLf &
                          "			LEFT OUTER JOIN  Clientes AS Deposito ON ApuracaoDeCustos.Deposito_Id = Deposito.Cliente_Id  " & vbCrLf &
                          "                        AND ApuracaoDeCustos.EndDeposito_Id = Deposito.Endereco_Id  " & vbCrLf &
                          "         LEFT OUTER JOIN  Clientes AS Destino ON ApuracaoDeCustos.EmpresaDestino_Id = Destino.Cliente_Id  " & vbCrLf &
                          "                        AND ApuracaoDeCustos.EndEmpresaDestino_Id = Destino.Endereco_Id   " & vbCrLf &
                          "			INNER JOIN  Produtos ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id " & vbCrLf &
                          "         Left JOIN  Produtos as Derivado ON ApuracaoDeCustos.ProdutoDerivado_Id = Derivado.Produto_Id " & vbCrLf &
                          "         INNER JOIN  PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id " & vbCrLf &
                          "         Left JOIN   TabelaDePrecosDeMercado ON ApuracaoDeCustos.Empresa_Id = TabelaDePrecosDeMercado.Empresa_Id " & vbCrLf &
                          "                        AND  ApuracaoDeCustos.EndEmpresa_Id = TabelaDePrecosDeMercado.EndEmpresa_Id  " & vbCrLf &
                          "                        AND  ApuracaoDeCustos.Deposito_Id = TabelaDePrecosDeMercado.Deposito_Id " & vbCrLf &
                          "                        AND  ApuracaoDeCustos.EndDeposito_Id = TabelaDePrecosDeMercado.EndDeposito_Id  " & vbCrLf &
                          "                        AND  ApuracaoDeCustos.Produto_Id = TabelaDePrecosDeMercado.Produto_Id " & vbCrLf &
                          "                        AND  month(Data_Id)  = " & DdlMes.SelectedValue & " AND year(Data_Id)= '" & DdlAno.SelectedValue & "' " & vbCrLf &
                          " WHERE     ApuracaoDeCustos.Ano_Id = '" & DdlAno.SelectedValue & "' AND ApuracaoDeCustos.Mes_Id = " & DdlMes.SelectedValue & "  " & vbCrLf

        If ddlEmpresa.Text <> "" Then
            sql &= " And ApuracaoDeCustos.Empresa_ID = '" & Empresa(0) & "' " & vbCrLf &
                   " And ApuracaoDeCustos.EndEmpresa_Id = " & Empresa(1)
        End If
        If DdlDeposito.Text <> "" Then
            sql &= " And ApuracaoDeCustos.Deposito_Id = '" & Deposito(0) & "' " & vbCrLf &
                   " And ApuracaoDeCustos.EndDeposito_Id = " & Deposito(1)
        End If

        'If DdlProduto.Text <> "" Then
        '    Sql &= " And ApuracaoDeCustos.Produto_Id = '" & DdlProduto.SelectedValue & "' " & vbCrLf
        'End If

        If ucSelecaoProduto.TemSelecionado Then
            sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo_Id", "Produtos.Produto_Id")(0)
        End If

        sql &= " Order by  " & vbCrLf &
               "" & vbCrLf &
               " ApuracaoDeCustos.Empresa_Id,  ApuracaoDeCustos.EndEmpresa_Id, " & vbCrLf &
               " ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf &
               " ApuracaoDeCustos.Produto_Id,  ApuracaoDeCustos.CodigoDeCusto_Id " & vbCrLf &
               " Delete #temp " & vbCrLf &
               "  Where Peso           = 0 " & vbCrLf &
               "    AND ValorUnitario  = 0 " & vbCrLf &
               "    AND ValorDeMercado = 0 " & vbCrLf &
               "    AND ValorDoFrete   = 0 " & vbCrLf &
               "    AND ValorTotal     = 0 " & vbCrLf
        Return sql

    End Function

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkAuditoria_CheckedChanged(sender As Object, e As EventArgs) Handles chkAuditoria.CheckedChanged
        If chkAuditoria.Checked Then
            rdNotas.Checked = True
            rdRazao.Checked = False
            rdProducao.Checked = False
        Else
            rdNotas.Checked = False
            rdRazao.Checked = False
            rdProducao.Checked = False

        End If
        rdNotas.Enabled = chkAuditoria.Checked
        rdRazao.Enabled = chkAuditoria.Checked
        rdProducao.Enabled = chkAuditoria.Checked
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CustoMedio")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class
