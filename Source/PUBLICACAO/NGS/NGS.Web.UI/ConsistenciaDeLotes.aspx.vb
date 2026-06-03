Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml

Public Class ConsistenciaDeLotes
    Inherits BasePage

    Private Sql As String
    Private ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ConsistenciaDeLotes", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                Limpar()
            Else
                MsgBox(Me.Page, "", eTitulo.Info)
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
    End Sub

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

            ds = getDataset()

            Dim dt As DataTable = New DataTable()

            dt = ds.Tables(0)

            dt.Columns.Remove("Reduzido")
            dt.Columns.Remove("NomeEmpresa")
            dt.Columns.Remove("CidadeEmpresa")
            dt.Columns.Remove("EstadoEmpresa")

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
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório De Consistencia De Lotes.")

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
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "CONSISTENCIA DE LOTES")
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

                    ' criando cabeçalho da planilha com os dados do dataset
                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:F" & rowIndex).AutoFilter = True

                    'formatando células numéricas
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

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
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = Format(Today, "01/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        txtLote.Text = String.Empty
        txtTitulo.Text = String.Empty

        gridConsistenciaDeLotes.DataSource = Nothing
        gridConsistenciaDeLotes.DataBind()
        LiberaEmpresa()
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresas()
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Unidade: " & ddlUnidade.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à: " & txtDataFinal.Text
        End If

        Return param
    End Function

    Function getDataset() As DataSet
        Sql = " SELECT   Razao.Empresa_Id As Empresa,Razao.Movimento_Id as Movimento, Razao.Lote_Id as Lote, Razao.Sequencia_Id as Sequencia, Razao.Titulo," & vbCrLf &
              "  Case isnull(PC.Conta_ID,'') " & vbCrLf &
              "       WHEN '' " & vbCrLf &
              "          THEN 'CONTA NAO CADASTRADA' " & vbCrLf &
              "          ELSE Razao.Conta_ID + ' - ' + PC.Titulo " & vbCrLf &
              "       END AS Conta, " & vbCrLf &
              "  Razao.Produto, " & vbCrLf &
              " Razao.Cliente_Id as Reduzido, Razao.Custo, " & vbCrLf &
              " Razao.Historico,  " & vbCrLf &
              " Case isnull(PC.Conta_ID,'')  " & vbCrLf &
              "        WHEN ''  " & vbCrLf &
              "           THEN 0  " & vbCrLf &
              "           ELSE Razao.DebitoOficial  " & vbCrLf &
              "        END AS Debito,   " & vbCrLf &
              "  Case isnull(PC.Conta_ID,'')  " & vbCrLf &
              "        WHEN ''  " & vbCrLf &
              "           THEN 0  " & vbCrLf &
              "           ELSE Razao.CreditoOficial  " & vbCrLf &
              "        END AS Credito, " & vbCrLf &
              "'D' As Saldo, " & vbCrLf &
              " Clientes.Nome As NomeEmpresa, Clientes.Cidade as CidadeEmpresa, Clientes.Estado as EstadoEmpresa " & vbCrLf &
              " FROM Razao " & vbCrLf &
              " LEFT JOIN PlanoDeContas PC " & vbCrLf &
              " ON Razao.Conta_Id = PC.Conta_Id " & vbCrLf &
              " INNER JOIN Clientes " & vbCrLf &
              " ON Razao.Empresa_Id    = Clientes.Cliente_Id " & vbCrLf &
              " AND Razao.EndEmpresa_Id = Clientes.Endereco_Id " & vbCrLf &
              " WHERE Razao.Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf &
              "   And '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

        If chkConsolidarEmpresa.Checked Then
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Sql &= " And Razao.Empresa_Id LIKE '" & Left(Empresa(0), 8) & "%'"
        Else
            Dim empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Sql &= " And Razao.Empresa_Id = '" & empresa(0) & "' And Razao.EndEmpresa_Id = " & empresa(1) & vbCrLf
        End If

        If txtLote.Text <> "" Then
            Sql &= " And Razao.Lote_Id = " & txtLote.Text & vbCrLf
        End If

        If txtTitulo.Text <> "" Then
            Sql &= " And Razao.Titulo = " & txtTitulo.Text & vbCrLf
        End If

        Sql &= " ORDER BY Razao.Empresa_Id, Razao.Movimento_Id, Razao.Lote_Id, Razao.Sequencia_Id"

        ds = Banco.ConsultaDataSet(Sql, "Consistencia")

        Dim Saldo As Decimal = 0
        For Each dr As DataRow In ds.Tables(0).Rows
            Saldo += dr("Debito") - dr("Credito")
            If dr("Saldo") = "A" Then
                dr("Debito") = 0
                dr("Credito") = 0
            End If
            If Saldo > 0 Then
                dr("Saldo") = Saldo.ToString("N2") & "-D"
            End If
            If Saldo < 0 Then
                dr("Saldo") = (Saldo * -1).ToString("N2") & "-C"
            End If
            If Saldo = 0 Then
                dr("Saldo") = Saldo.ToString("N2") & "DC"
            End If

            If dr("Empresa").ToString.Length > 0 Then dr("Empresa") = Funcoes.FormatarCpfCnpj(dr("Empresa"))
            If dr("Reduzido").ToString.Length > 0 Then dr("Reduzido") = Funcoes.FormatarCpfCnpj(dr("Reduzido"))
        Next

        Return ds
    End Function

    Protected Sub LinkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles LinkConsultar.Click
        Try
            ds = getDataset()
            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Período sem movimento")
            Else
                gridConsistenciaDeLotes.DataSource = ds
                gridConsistenciaDeLotes.DataBind()

                Dim i As Integer = 0

                While i < gridConsistenciaDeLotes.Rows.Count
                    If gridConsistenciaDeLotes.Rows(i).Cells(5).Text = "CONTA NAO CADASTRADA" Then
                        gridConsistenciaDeLotes.Rows(i).ForeColor = Color.Red

                    End If
                    i += 1
                End While


            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString))
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ConsistenciaDeLotes", "RELATORIO") Then
                ds = getDataset()

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", IIf(ddlEmpresa.SelectedIndex > 0, "Relatório De Consistencia De Lotes.", "Relatório De Consistencia De Lotes Por Empresa."))
                parameters.Add("Parametros", getParam())

                If ddlEmpresa.SelectedIndex > 0 Then
                    Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                    Funcoes.BindReport(Me.Page, ds, IIf(ddlEmpresa.SelectedIndex > 0, "Cr_ConsistenciasDeLotes", "Cr_ConsistenciasDeLotesPorEmpresa"), eExportType.PDF, parameters, False, "", Empresa(0), Empresa(1))
                Else
                    Funcoes.BindReport(Me.Page, ds, IIf(ddlEmpresa.SelectedIndex > 0, "Cr_ConsistenciasDeLotes", "Cr_ConsistenciasDeLotesPorEmpresa"), eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para tirar Relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorioDados() 'Excel Dados
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConsistenciaDeLotes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class