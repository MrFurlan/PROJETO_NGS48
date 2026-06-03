Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style

Public Class TotalizaLotes
    Inherits BasePage

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TotalizaLotes", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                txtDataInicial.Text = Format(Today, "01/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LinkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles LinkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("TotalizaLotes", "LEITURA") Then
                If ValidarCampos() Then
                    Dim strSQL As String
                    Dim ds As New DataSet
                    Dim Empresa As String = ddlEmpresa.SelectedValue.Split("-")(0)
                    Dim EndEmpresa As String = ddlEmpresa.SelectedValue.Split("-")(1)
                    Dim Conta As String = HttpContext.Current.Session("Conta")
                    Dim EndCliente As String = HttpContext.Current.Session("ClienteeND")
                    Dim Custo As String = ""
                    Dim Produto As String = ""

                    If EndCliente = "" Then
                        EndCliente = 0
                    End If

                    If Custo = "" Then
                        Custo = 0
                    End If

                    If Produto = "" Then
                        Produto = 0
                    End If

                    strSQL = "SELECT Movimento_Id as Movimento, Lote_Id as Lote,  Sum(DebitoOficial) as DebitoOficial, Sum(CreditoOficial) as CreditoOficial, " & vbCrLf & _
                             "SUM(DebitoOficial - CreditoOficial) AS SaldoOficial, Sum(DebitoMoeda) as DebitoMoeda, " & vbCrLf & _
                             "Sum(CreditoMoeda) as CreditoMoeda, SUM(DebitoMoeda - CreditoMoeda) AS SaldoMoeda, 'D' as Saldo " & vbCrLf & _
                             "FROM Razao WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & vbCrLf

                    If txtLote.Text <> "" Then
                        strSQL &= "And Lote_Id = " & txtLote.Text
                    End If

                    strSQL &= " And Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                              "GROUP BY Movimento_Id, Lote_Id " & vbCrLf & _
                              "ORDER BY Movimento_Id, Lote_Id" & vbCrLf

                    Try
                        ds = Banco.ConsultaDataSet(strSQL, "Razao")

                        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                            Dim dra As DataRow
                            Dim Saldo As Decimal = 0

                            For Each dra In ds.Tables(0).Rows
                                Saldo += dra("DebitoOficial") - dra("CreditoOficial")
                                If dra("Saldo") = "A" Then
                                    dra("DebitoOficial") = 0
                                    dra("CreditoOficial") = 0
                                End If
                                If Saldo > 0 Then
                                    dra("Saldo") = Saldo.ToString("N2") & "-D"
                                End If
                                If Saldo < 0 Then
                                    dra("Saldo") = (Saldo * -1).ToString("N2") & "-C"
                                End If
                                If Saldo = 0 Then
                                    dra("Saldo") = Saldo.ToString("N2") & "DC"
                                End If
                            Next
                            GridView1.DataSource = ds
                            GridView1.DataBind()
                        Else
                            MsgBox(Me.Page, "Período sem movimento.")
                        End If
                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                    End Try
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("TotalizaLotes", "RELATORIO") Then
                If ValidarCampos() Then
                    Try
                        Dim strParam As String = ""
                        Dim ds As DataSet = getDataSet()
                        Dim dra As DataRow
                        Dim Saldo As Decimal = 0

                        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count = 2 AndAlso ds.Tables(1).Rows.Count > 0 Then
                            For Each dra In ds.Tables(1).Rows
                                Saldo += dra("DebitoOficial") - dra("CreditoOficial")

                                If Saldo > 0 Then
                                    dra("Saldo") = Saldo.ToString("N2") & "-D"
                                End If
                                If Saldo < 0 Then
                                    dra("Saldo") = (Saldo * -1).ToString("N2") & "-C"
                                End If
                                If Saldo = 0 Then
                                    dra("Saldo") = Saldo.ToString("N2") & "DC"
                                End If
                            Next

                            For Each row As DataRow In ds.Tables(0).Rows
                                strParam = row("Nome") & "," & row("Cidade") & " - " & row("Estado")
                            Next

                            Dim param As New Dictionary(Of String, Object)()
                            param.Add("EmpresaCabecalho", strParam.Split(",")(0))
                            param.Add("CidadeCabecalho", strParam.Split(",")(1))
                            param.Add("Periodo", "Período.: " & txtDataInicial.Text & " A " & txtDataFinal.Text)

                            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

                            Funcoes.BindReport(Me.Page, ds, "Cr_TotalizadorDeLotes", eExportType.PDF, param, False, "", Empresa(0), Empresa(1))
                        Else
                            MsgBox(Me.Page, "Período sem movimento.")
                        End If
                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                    End Try
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorioDados() 'Excel Dados
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            txtDataInicial.Text = Format(Today, "01/01/yyyy")
            txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

            GridView1.DataSource = New List(Of Object)
            GridView1.DataBind()
            LiberaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TotalizaLotes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.Text) Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresa.Text) Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf chkConsolidarEmpresa.Checked Then
            MsgBox(Me.Page, "Consolidar empresa indisponível para esse processo.")
            Return False
        End If
        Return True
    End Function

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub VerificaUnidade()
        Dim Sql As String = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
                            "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
                            "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
                            " from Usuarios" & vbCrLf & _
                            " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Private Function getDataSet() As DataSet
        Dim ds As New DataSet
        Dim sql As String = "  SELECT Nome, Cidade, Estado " & vbCrLf &
                            "   FROM   Clientes" & vbCrLf &
                            "   Where Cliente_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' and Endereco_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Clientes")

        sql = "SELECT Empresa_Id, EndEmpresa_Id, Movimento_Id AS Movimento, REPLICATE('0', 4 - LEN(CAST(Lote_Id AS varchar))) + CAST(Lote_Id AS varchar) as Lote, " & vbCrLf &
              "Sum(DebitoOficial) as DebitoOficial, Sum(CreditoOficial) as CreditoOficial, " & vbCrLf &
              "SUM(DebitoOficial - CreditoOficial) AS SaldoOficial, Sum(DebitoMoeda) as DebitoMoeda, " & vbCrLf &
              "Sum(CreditoMoeda) as CreditoMoeda, SUM(DebitoMoeda - CreditoMoeda) AS SaldoMoeda,'D' as Saldo " & vbCrLf &
              "FROM Razao " & vbCrLf &
              " WHERE Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

        If Not chkConsolidarEmpresa.Checked Then
            sql &= "AND Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' And EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtLote.Text) Then
            sql &= "AND Lote_Id = " & txtLote.Text
        End If

        sql &= "   GROUP BY Empresa_Id, EndEmpresa_Id, Movimento_Id, Lote_Id " & vbCrLf &
               "   ORDER BY Empresa_Id, EndEmpresa_Id, Movimento_Id, Lote_Id" & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(sql, "TotalizadorDeLotes"))

        Return ds
    End Function

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub EmitirRelatorioDados()
        Try
            Dim Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
            Dim data As String = CStr(txtDataInicial.Text & " - " & txtDataFinal.Text)
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


                    'criando planilha e título
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Totalizador de Lotes.")

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
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "TOTALIZADOR DE LOTES")
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

                    ' criando cabeçalho da planilha com os dados do dataset
                    For Each col As DataColumn In ds.Tables(1).Columns
                        If col.ColumnName <> "Row" Then
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        End If
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:I" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(1).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(1).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células valores
                        worksheet.Cells(String.Format("D{0}:I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count)
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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region


End Class