Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class ExtratoDeConta
    Inherits BasePage

    Dim Sql As String
    Private ds As DataSet
    Private Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Contabil)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ExtratoDeConta", "ACESSAR") Then
                    Limpar()
                    CargaUnidade()
                    VerificaUnidade()
                    CargaCentroDeCusto()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objClienteEx" & HID.Value) IsNot Nothing Then
                Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteEx" & HID.Value), [Lib].Negocio.Cliente)
                hdfCliente.Value = objCliente.Codigo & " - " & objCliente.CodigoEndereco
                txtCliente.Text = objCliente.Codigo & " - " & objCliente.CodigoEndereco & " ..." & objCliente.Nome & " / " & objCliente.Cidade & "-" & objCliente.CodigoEstado
                Session.Remove("objClienteEx" & HID.Value)
            ElseIf Session("objContaEx" & HID.Value) IsNot Nothing Then
                Dim objConta As [Lib].Negocio.PlanoDeConta = CType(Session("objContaEx" & HID.Value), [Lib].Negocio.PlanoDeConta)
                hdfConta.Value = objConta.Conta
                txtConta.Text = objConta.Conta & " - " & objConta.Titulo
                If Not objConta.TemCliente Then
                    hdfCliente.Value = ""
                    txtCliente.Text = ""
                End If
                Session.Remove("objContaEx" & HID.Value)
                CarregarGrupos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
    End Sub

    Private Sub CargaCentroDeCusto()
        ddl.Carregar(DdlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "Len(CentroDeCusto_Id) = 5", True)
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresas()
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub CargaProdutos()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo= " & ddlGrupo.SelectedValue & "", True)
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = Format(Today, "01/01/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        Session.Remove("objClienteEx" & HID.Value)
        Session.Remove("objContaEx" & HID.Value)
        hdfConta.Value = ""
        txtConta.Text = ""
        hdfCliente.Value = ""
        txtCliente.Text = ""
        DdlCentroDeCusto.SelectedIndex = 0
        ddlGrupo.Items.Clear()
        ddlProduto.Items.Clear()
        chkContasDoPeriodo.Checked = False
        chkHistorico.Checked = False
        ckPorCliente.Checked = False
        HID.Value = Guid.NewGuid().ToString
        ucConsultaPlanoDeContas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Function getDataSet(ByVal tipo As String) As DataSet
        Dim strSql As String
        Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim strCliente() As String = hdfCliente.Value.ToString.Split("-")
        Dim objConta As New [Lib].Negocio.PlanoDeConta("", 0, hdfConta.Value)
        Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
        Dim objCliente As [Lib].Negocio.Cliente = Nothing

        If objConta IsNot Nothing AndAlso objConta.TemCliente Then
            objCliente = New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
        End If

        Dim Custo As String = ""
        Dim Produto As String = ""
        If ddlProduto.SelectedIndex > 0 Then
            Produto = ddlProduto.SelectedValue
        ElseIf ddlGrupo.SelectedIndex > 0 Then
            Produto = ddlGrupo.SelectedValue
        End If

        If Custo = "" Then
            Custo = 0
        End If

        If Produto = "" Then
            Produto = 0
        End If

        Dim DataInicial As String = Format(CDate(txtDataInicial.Text), "yyyy/MM/dd")
        Dim DataFinal As String = Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")
        Dim PDataInicial As String = Format(CDate(txtDataInicial.Text), "dd/MM/yyyy")
        Dim PDataFinal As String = Format(CDate(txtDataFinal.Text), "dd/MM/yyyy")
        Dim TipoDeRelatorio As String = "PDF"

        strSql = "  SELECT     Movimento_Id as Movimento, Lote_Id as Lote, " & vbCrLf & _
                 " Sequencia_Id as Sequencia, " & vbCrLf & _
                 " Titulo, " & vbCrLf & _
                 " case " & _
                 "	when len(isnull(Produto, '')) > 0" & _
                 "		then Produto" & _
                 "		else" & _
                 "			case" & _
                 "				when len(isnull(Produto_NF, '')) > 0" & _
                 "					then Produto_NF" & _
                 "					else ''" & _
                 "			end" & _
                 "	end as Produto, " & _
                 " Custo, Historico, DebitoOficial As Debito, CreditoOficial as Credito, " & vbCrLf

        If tipo = "R" Then
            strSql &= "convert(Decimal(15,2),0) as Saldo, " & vbCrLf
        Else
            strSql &= "'D' AS Saldo, " & vbCrLf
        End If
        strSql &= "2 as Ordem  " & vbCrLf
        strSql &= " into #temp " & vbCrLf
        strSql &= " FROM Razao" & vbCrLf
        strSql &= " WHERE Empresa_Id = '" & objEmpresa.Codigo & "' And EndEmpresa_Id = " & objEmpresa.CodigoEndereco & " And Conta_Id = '" & objConta.Conta & "' " & vbCrLf

        If objCliente IsNot Nothing AndAlso objCliente.Codigo.Length > 0 Then
            strSql &= " And  Cliente_Id = '" & objCliente.Codigo & "' And EndCliente_Id = " & objCliente.CodigoEndereco & vbCrLf
        End If

        If Custo <> "0" Then
            strSql &= " And Custo = " & Custo & vbCrLf
        End If
        If Produto <> "0" Then
            strSql &= " And (Produto LIKE '" & Produto & "%' OR Produto_NF LIKE '" & Produto & "%') " & vbCrLf
        End If
        strSql &= " And Movimento_Id BETWEEN '" & DataInicial & "' AND '" & DataFinal & "'" & vbCrLf

        strSql &= " update #temp set Saldo = Debito - Credito " & vbCrLf

        strSql &= " insert into #temp " & vbCrLf
        strSql &= "SELECT '" & CDate(DataInicial).AddDays(-1).ToString("yyyy/MM/dd") & "' as Movimento, ' ' as Lote, ' ' as Sequencia, ' ' as Titulo, ' ' as Produto, '0' as Custo, 'SALDO INICIAL...........' as Historico, 0.00 as Debito, 0.00 as Credito, " & vbCrLf
        If tipo = "R" Then
            strSql &= "COALESCE(SUM(DebitoOficial - CreditoOficial), 0) AS Saldo, " & vbCrLf
        Else
            strSql &= "'A' AS Saldo, " & vbCrLf
        End If
        strSql &= "1 as Ordem  " & vbCrLf
        strSql &= " FROM Razao " & vbCrLf
        strSql &= " WHERE Empresa_Id = '" & objEmpresa.Codigo & "' And EndEmpresa_Id = " & objEmpresa.CodigoEndereco & " And Conta_Id = '" & objConta.Conta & "' " & vbCrLf

        If objCliente IsNot Nothing AndAlso objCliente.Codigo.Length > 0 Then
            strSql &= " And  Cliente_Id = '" & objCliente.Codigo & "' And EndCliente_Id = " & objCliente.CodigoEndereco & vbCrLf
        End If

        If Custo <> "0" Then
            strSql &= " And Custo = " & Custo & vbCrLf
        End If
        If Produto <> "0" Then
            strSql &= " And (Produto LIKE '" & Produto & "%' OR Produto_NF LIKE '" & Produto & "%') " & vbCrLf
        End If
        strSql &= " And Movimento_Id < '" & DataInicial & "'" & vbCrLf

        strSql &= " SELECT   Movimento, Lote, Sequencia, Titulo, Produto, Custo, Historico, Debito, Credito, Saldo, Ordem " & vbCrLf
        strSql &= " FROM #Temp " & vbCrLf

        If chkHistorico.Checked = True Then
            strSql &= " ORDER BY Ordem, Historico" & vbCrLf
        Else
            strSql &= " ORDER BY Movimento, Lote, Sequencia" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(strSql, "Extrato")
        Return ds
    End Function

    Function Validar() As Boolean
        Dim objConta As New [Lib].Negocio.PlanoDeConta("", 0, hdfConta.Value)
        If ddlUnidade.Text = "" Then
            Mensagem = "Unidade de negócio é obrigatório..."
            Return False
        ElseIf ddlEmpresa.Text = "" Then
            Mensagem = "Empresa é obrigatório..."
            Return False
        ElseIf txtConta.Text = "" Then
            Mensagem = "Conta é obrigatório..."
            Return False
        ElseIf objConta IsNot Nothing AndAlso objConta.TemCliente AndAlso txtCliente.Text = "" Then
            Mensagem = "Cliente é obrigatório..."
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub ddlGrupo_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaProdutos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarGrupos()
        Dim objConta As New [Lib].Negocio.PlanoDeConta("", 0, hdfConta.Value)
        If objConta IsNot Nothing AndAlso objConta.Conta.Length > 0 Then
            If objConta.TemProduto Then
                ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
            End If
        End If
    End Sub

    Protected Sub btnConsultaConta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultaConta.Click
        Try
            Dim txtConta As TextBox = ucConsultaPlanoDeContas.FindControl("txtConta")
            If txtConta IsNot Nothing Then
                txtConta.Text = String.Empty
                txtConta.Focus()
            End If
            Session("TipoConta" & HID.Value) = "Credito"
            Session("Codigo" & HID.Value) = hdfConta.ClientID
            Session("Descricao" & HID.Value) = txtConta.ClientID
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me.Page, "objContaEx" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCliente.Click
        Try
            If Not ckPorCliente.Checked Then
                Dim objConta As New [Lib].Negocio.PlanoDeConta("", 0, hdfConta.Value)
                If objConta IsNot Nothing AndAlso objConta.TemCliente Then
                    ucConsultaClientes.Limpar()
                    Popup.ConsultaDeClientes(Me.Page, "objClienteEx" & HID.Value, "txtNome")
                Else
                    MsgBox(Me.Page, "Conta não permite cliente!")
                End If
            Else
                ucConsultaClientes.Limpar()
                Popup.ConsultaDeClientes(Me.Page, "objClienteEx" & HID.Value, "txtNome")
                txtConta.Text = ""
                hdfConta.Value = ""
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ExtratoDeConta")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Session.Remove("objClienteEx" & HID.Value)
            Session.Remove("objContaEx" & HID.Value)

            hdfConta.Value = ""
            txtConta.Text = ""
            hdfCliente.Value = ""
            txtCliente.Text = ""
            DdlCentroDeCusto.SelectedIndex = 0
            ddlGrupo.Items.Clear()
            ddlProduto.Items.Clear()

            txtDataInicial.Text = Format(Today, "01/01/yyyy")
            txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

            HID.Value = Guid.NewGuid().ToString
            ucConsultaPlanoDeContas.SetarHID(HID.Value)
            ucConsultaClientes.SetarHID(HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            If chkContasDoPeriodo.Checked Then
                EmitirRelatorioGeral(True)
            Else
                EmitirRelatorio(True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            If chkContasDoPeriodo.Checked Then
                EmitirRelatorioGeral(False)
            Else
                EmitirRelatorio(False)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Pdf Then

                Dim crpt As New ReportDocument()

                Try
                    If Funcoes.VerificaPermissao("ExtratoDeConta", "RELATORIO") Then
                        If Validar() Then
                            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
                            Dim objConta As New [Lib].Negocio.PlanoDeConta("", 0, hdfConta.Value)
                            Dim DataInicial As String = Format(CDate(txtDataInicial.Text), "yyyy/MM/dd")
                            Dim DataFinal As String = Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")
                            Dim PDataInicial As String = Format(CDate(txtDataInicial.Text), "dd/MM/yyyy")
                            Dim PDataFinal As String = Format(CDate(txtDataFinal.Text), "dd/MM/yyyy")

                            Dim dsRel As DataSet = getDataSet("R")
                            If dsRel Is Nothing OrElse dsRel.Tables(0).Rows.Count = 0 Then
                                MsgBox(Me.Page, "Período sem movimento.")
                            Else
                                Dim parameters = New Dictionary(Of String, Object)()
                                parameters.Add("Periodo", "Período.: " & PDataInicial & " A " & PDataFinal)
                                If objConta IsNot Nothing AndAlso objConta.TemCliente Then
                                    Dim strCliente() As String = hdfCliente.Value.ToString.Split("-")
                                    Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
                                    parameters.Add("Conta", objConta.Conta & Funcoes.FormatarCpfCnpj(objCliente.Codigo) & " - " & objCliente.Nome)
                                Else
                                    parameters.Add("Conta", objConta.Conta & " - " & objConta.Titulo)
                                End If
                                Funcoes.BindReport(Me.Page, ds, "Cr_ExtratoDeContas", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters, False, "", strEmpresa(0), strEmpresa(1))
                            End If
                        Else
                            MsgBox(Me.Page, Mensagem)
                        End If
                    Else
                        MsgBox(Me.Page, "Usuário sem permissão para emitir relatório!")
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                Finally
                    crpt.Close()
                    crpt.Dispose()
                End Try
            ElseIf ckPorCliente.Checked Then
                Try
                    Dim ds As DataSet = getDataSetExtratoPorCliente()

                    Dim colunas As New Dictionary(Of String, eTipoCampo)
                    colunas.Add("Movimento", eTipoCampo.Data)
                    colunas.Add("DebitoOficial", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("CreditoOficial", eTipoCampo.ValorSemTotalizador)

                    Funcoes.BindExcelOffice(Me.Page, ds, "Extrato por Cliente", colunas)
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            Else
                If validaCamposExcel() Then
                    Dim ds As DataSet = getDataSetExcel()

                    If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        Throw New Exception("Nenhum resultado encontrado!")
                    Else
                        BindExcelOffice(ds, "Extrato de Conta")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorioGeral(ByVal Pdf As Boolean)
        If ddlUnidade.Text = "" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório...", eTitulo.Info)
        ElseIf ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório...", eTitulo.Info)
        Else

            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim DataInicial As String = Format(CDate(txtDataInicial.Text), "yyyy/MM/dd")
            Dim DataFinal As String = Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")
            Dim strSql As String


            strSql = " SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id " & vbCrLf & _
                     " WHERE Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' And EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & vbCrLf & _
                     "   And Movimento_Id BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy/MM/dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy/MM/dd") & "'" & vbCrLf & _
                     " GROUP BY Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id" & vbCrLf & _
                     " ORDER BY Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id"

            ds = Banco.ConsultaDataSet(strSql, "ContasDoExtrato")



        End If
    End Sub

    Private Function validaCamposExcel()
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a Empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(hdfConta.Value) Then
            MsgBox(Me.Page, "Informe a Conta.")
            Return False
        End If
        Return True
    End Function

    Public Sub BindExcelOffice(ByVal ds As DataSet, ByVal TituloAba As String)
        If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                'emitir excel.xsls do office / relatório padrão em lista
                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)

                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                        Sql = "SELECT Top 1 Cliente_Id as CNPJ, Nome, Cidade, Estado" & vbCrLf & _
                              "  FROM Clientes" & vbCrLf & _
                              " where Cliente_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' "

                        ds.Merge(Banco.ConsultaDataSet(Sql, "Empresas"))

                        worksheet.Cells(rowIndex, columnIndex).Value = ds.Tables("Empresas").Rows(0)("Nome")
                        worksheet.Cells(rowIndex, columnIndex, rowIndex, 8).Merge = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        rowIndex += 1
                        columnIndex += 1

                        worksheet.Cells(rowIndex, columnIndex).Value = ds.Tables("Empresas").Rows(0)("Cidade") & " - " & ds.Tables("Empresas").Rows(0)("Estado")
                        rowIndex += 1

                        worksheet.Cells(rowIndex, columnIndex).Value = "CNPJ" & Funcoes.FormatarCpfCnpj(ds.Tables("Empresas").Rows(0)("Cnpj"))
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        rowIndex += 1

                        worksheet.Cells(rowIndex, columnIndex).Value = "Extrato de Contas"
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        columnIndex = 1
                        rowIndex = 6

                        If Not String.IsNullOrWhiteSpace(hdfConta.Value) Then
                            worksheet.Cells(rowIndex, columnIndex).Value = txtConta.Text
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex, rowIndex, 4).Merge = True
                        Else
                            worksheet.Cells(rowIndex, columnIndex).Value = ddlGrupo.SelectedItem.Text
                            If Not String.IsNullOrWhiteSpace(hdfCliente.Value) Then
                                columnIndex += 1
                                worksheet.Cells(rowIndex, columnIndex).Value = Funcoes.FormatarCpfCnpj(hdfCliente.Value)
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            End If
                        End If
                        rowIndex = 8
                        columnIndex = 5

                        worksheet.Cells(rowIndex, columnIndex).Value = "Emissão : " & Now
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex, rowIndex, 8).Merge = True
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                        columnIndex = 1

                        worksheet.Cells(rowIndex, columnIndex).Value = "Período " & txtDataInicial.Text & " À " & txtDataFinal.Text
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex, rowIndex, 4).Merge = True

                        rowIndex += 1

                        'aplicando formatação nas células do Cabecalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, 10)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        worksheet.Cells(rowIndex, columnIndex).Value = "Empresa"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Data"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Lote"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Sequência"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Titulo"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Produto"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Histórico"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Débitos"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Créditos"
                        columnIndex += 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "Saldo"
                        rowIndex += 1

                        Dim saldo As Decimal = ds.Tables(0).Rows(0)(18)

                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If col.ColumnName = "Empresa" OrElse col.ColumnName = "Data" OrElse col.ColumnName = "Lote" OrElse col.ColumnName = "Sequencia" OrElse col.ColumnName = "Titulo" OrElse col.ColumnName = "Produto" OrElse col.ColumnName = "Historico" OrElse
                                   col.ColumnName = "Debitos" OrElse col.ColumnName = "Creditos" OrElse col.ColumnName = "Saldo" Then
                                    If col.ColumnName = "Saldo" Then
                                        saldo = saldo + row("Debitos") - row("Creditos")
                                        worksheet.Cells(rowIndex, columnIndex).Value = saldo

                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)

                                    End If
                                    If IsDate(row(col.ColumnName)) Then
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                    ElseIf IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString().Contains(",") Then
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    ElseIf IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString().Contains(",") Then
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                    End If
                                    columnIndex += 1
                                End If
                            Next

                            'formatações de celulas
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, 10)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'aplicando formatação nas células do rodapé
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, 10)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        worksheet.Cells(9, 1, 9, 8).AutoFilter = True

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linham
                        worksheet.View.FreezePanes(10, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

    Private Function getDataSetExtratoPorCliente() As DataSet
        Dim ds As New DataSet
        Try
            If FinanceiroNovo Then
                Sql = "  SELECT   Empresa.Reduzido AS Empresa, Empresa.Nome AS NomeDaEmpresa, Empresa.Cidade AS CidadeDaEmpresa, Razao.Conta_Id as Conta,         " & vbCrLf & _
                  "           PlanoDeContas.Titulo AS TituloDaConta, Razao.Cliente_Id, Clientes.Nome AS NomeDoCliente, Razao.Movimento_Id as Movimento,       " & vbCrLf & _
                  "           Razao.Lote_Id as Lote, Razao.Sequencia_Id as Sequencia, Razao.Titulo,                                                           " & vbCrLf & _
                  "           CASE WHEN Razao.Pedido > 0                                                                                                      " & vbCrLf & _
                  "				THEN Razao.Pedido                                                                                                           " & vbCrLf & _
                  "				ELSE ContasAReceber.Pedido                                                                                                  " & vbCrLf & _
                  "		   END AS Pedido,                                                                                                                   " & vbCrLf & _
                  "		   ISNULL(Razao.Numero_Nf, '') AS Nota, ISNULL(Razao.Cfop_NF, '') AS CFOP, ISNULL(Cfop.Descricao, '') AS DescricaoCFOP,             " & vbCrLf & _
                  "                        Razao.Historico, Razao.DebitoOficial, Razao.CreditoOficial, Razao.Serie_Nf, Razao.Numero_Nf, NotasFiscais.Operacao," & vbCrLf & _
                  "                        NotasFiscais.SubOperacao, SubOperacoes.Descricao     " & vbCrLf & _
                  "                        FROM(NotaFiscalXTitulo)                              " & vbCrLf & _
                  "		RIGHT OUTER JOIN SubOperacoes                                         " & vbCrLf & _
                  "		INNER JOIN NotasFiscais                                               " & vbCrLf & _
                  "			ON SubOperacoes.Operacao_Id = NotasFiscais.Operacao               " & vbCrLf & _
                  "			AND SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao       " & vbCrLf & _
                  "		RIGHT OUTER JOIN Razao                                                " & vbCrLf & _
                  "        INNER JOIN Clientes AS Empresa                                       " & vbCrLf & _
                  "			ON Razao.Empresa_Id = Empresa.Cliente_Id                          " & vbCrLf & _
                  "			AND Razao.EndEmpresa_Id = Empresa.Endereco_Id                     " & vbCrLf & _
                  "     INNER JOIN PlanoDeContas                                             " & vbCrLf & _
                  "			ON Razao.Conta_Id = PlanoDeContas.Conta_Id                        " & vbCrLf & _
                  "     INNER JOIN Clientes                                                  " & vbCrLf & _
                  "			ON Razao.Cliente_Id = Clientes.Cliente_Id                         " & vbCrLf & _
                  "			AND Razao.EndCliente_Id = Clientes.Endereco_Id                    " & vbCrLf & _
                  "         ON NotasFiscais.Empresa_Id = Razao.Empresa_Id                    " & vbCrLf & _
                  "       	AND NotasFiscais.EndEmpresa_Id = Razao.EndEmpresa_Id    " & vbCrLf & _
                  "       	AND NotasFiscais.Serie_Id = Razao.Serie_Nf              " & vbCrLf & _
                  "       	AND NotasFiscais.Nota_Id = Razao.Numero_Nf              " & vbCrLf & _
                  "       	AND NotasFiscais.EntradaSaida_Id = Razao.EntradaSaida_Nf" & vbCrLf & _
                  "       	AND NotasFiscais.Cliente_Id = Razao.Cliente_Nf          " & vbCrLf & _
                  "       	AND NotasFiscais.EndCliente_Id = Razao.EndCliente_Nf    " & vbCrLf & _
                  "     LEFT OUTER JOIN Cfop                                           " & vbCrLf & _
                  "			ON Razao.Cfop_NF = Cfop.Cfop_Id                             " & vbCrLf & _
                  "     LEFT OUTER JOIN Titulos ContasAReceber                         " & vbCrLf & _
                  "			ON Razao.Titulo = ContasAReceber.Titulo_Id                  " & vbCrLf & _
                  "			AND RAZAO.Conta_Id = ContasAReceber.ContaContabilCliFor     " & vbCrLf & _
                  "			AND ContasAReceber.RecPag = 'R'                             " & vbCrLf & _
                  "			ON NotaFiscalXTitulo.Titulo_Id = Razao.Titulo               " & vbCrLf & _
                  "                        WHERE 1 = 1                                 " & vbCrLf
            Else
                Sql = " SELECT    Empresa.Reduzido AS Empresa, Empresa.Nome AS NomeDaEmpresa, Empresa.Cidade AS CidadeDaEmpresa, Razao.Conta_Id as Conta," & vbCrLf & _
                      "           PlanoDeContas.Titulo AS TituloDaConta, Razao.Cliente_Id, Clientes.Nome AS NomeDoCliente, Razao.Movimento_Id as Movimento, Razao.Lote_Id as Lote, Razao.Sequencia_Id as Sequencia," & vbCrLf & _
                      "           Razao.Titulo, CASE WHEN Razao.Pedido > 0 THEN Razao.Pedido ELSE ContasAReceber.Pedido END AS Pedido, ISNULL(Razao.Numero_Nf, '') AS Nota," & vbCrLf & _
                      "           ISNULL(Razao.Cfop_NF, '') AS CFOP, ISNULL(Cfop.Descricao, '') AS DescricaoCFOP, Razao.Historico, Razao.DebitoOficial, Razao.CreditoOficial, Razao.Serie_Nf," & vbCrLf & _
                      "           Razao.Numero_Nf , NotasFiscais.Operacao, NotasFiscais.SubOperacao, SubOperacoes.Descricao" & vbCrLf & _
                      " FROM  NotaFiscalXTitulo RIGHT OUTER JOIN" & vbCrLf & _
                      "           SubOperacoes INNER JOIN" & vbCrLf & _
                      "           NotasFiscais ON SubOperacoes.Operacao_Id = NotasFiscais.Operacao AND SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao RIGHT OUTER JOIN" & vbCrLf & _
                      "           Razao INNER JOIN" & vbCrLf & _
                      "           Clientes AS Empresa ON Razao.Empresa_Id = Empresa.Cliente_Id AND Razao.EndEmpresa_Id = Empresa.Endereco_Id INNER JOIN" & vbCrLf & _
                      "           PlanoDeContas ON Razao.Conta_Id = PlanoDeContas.Conta_Id INNER JOIN" & vbCrLf & _
                      "           Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id AND Razao.EndCliente_Id = Clientes.Endereco_Id ON NotasFiscais.Empresa_Id = Razao.Empresa_Id AND" & vbCrLf & _
                      "           NotasFiscais.EndEmpresa_Id = Razao.EndEmpresa_Id AND NotasFiscais.Serie_Id = Razao.Serie_Nf AND NotasFiscais.Nota_Id = Razao.Numero_Nf AND" & vbCrLf & _
                      "           NotasFiscais.EntradaSaida_Id = Razao.EntradaSaida_Nf AND NotasFiscais.Cliente_Id = Razao.Cliente_Nf AND" & vbCrLf & _
                      "           NotasFiscais.EndCliente_Id = Razao.EndCliente_Nf LEFT OUTER JOIN" & vbCrLf & _
                      "           Cfop ON Razao.Cfop_NF = Cfop.Cfop_Id LEFT OUTER JOIN" & vbCrLf & _
                      "           ContasAReceber ON Razao.Titulo = ContasAReceber.Registro_Id ON NotaFiscalXTitulo.Titulo_Id = Razao.Titulo" & vbCrLf & _
                      " WHERE 1=1" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(hdfCliente.Value) Then
                Sql &= "And Razao.Cliente_Id = '" & hdfCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                    "   And Razao.EndCliente_Id = " & hdfCliente.Value.Split("-")(1) & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                Sql &= " And  (Razao.Movimento_Id  between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') " & vbCrLf
            End If

            Sql &= " ORDER BY Empresa.Reduzido, Razao.Conta_Id, Razao.Movimento_Id"

            ds = Banco.ConsultaDataSet(Sql, "ExtratodeContasPorCliente")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return ds
    End Function

    Private Function getDataSetExcel() As DataSet
        Try
            Dim sql As String = ""
            Dim Cliente As String = ""

            sql = " SELECT          Empresa.Reduzido AS RedEmpresa, Razao.Empresa_Id AS Empresa, Razao.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa," & vbCrLf &
                  "           Empresa.Cidade AS CidadeEmpresa, Razao.Conta_Id AS Conta, PlanoDeContas.Titulo AS NomeConta, Razao.Cliente_Id AS Cliente," & vbCrLf &
                  "           Razao.EndCliente_Id AS EndCliente, '' AS NomeCliente, '' AS CidadeCliente, '" & txtDataInicial.Text.ToSqlDate() & "' AS Data, 0 AS Lote, 0 AS Sequencia, 0 as Titulo, 0 as Produto, " & vbCrLf &
                  "           'Saldo Anterior.........' AS Historico, 0 AS Debitos, 0 AS Creditos, SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS Saldo" & vbCrLf &
                  "  FROM     GruposDeEstoques INNER JOIN" & vbCrLf &
                  "           Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo RIGHT OUTER JOIN" & vbCrLf &
                  "           Razao INNER JOIN" & vbCrLf &
                  "           Clientes AS Empresa ON Razao.Empresa_Id = Empresa.Cliente_Id AND Razao.EndEmpresa_Id = Empresa.Endereco_Id ON" & vbCrLf &
                  "           Produtos.Produto_Id = Razao.Produto_NF LEFT OUTER JOIN" & vbCrLf &
                  "           PlanoDeContas ON Razao.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf &
                  " WHERE     (Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "') " & vbCrLf

            If hdfConta.Value <> "" Then
                sql &= " And Razao.Conta_Id = '" & Left(hdfConta.Value, 9) & "'" & vbCrLf
            End If

            If ddlGrupo.SelectedValue <> "" And hdfConta.Value = "" And hdfCliente.Value = "" Then
                sql &= " And Razao.Conta_Id Like '" & Left(ddlGrupo.SelectedValue, 7) & "%'" & vbCrLf
            End If

            'If hdfCliente.Value <> "" Then
            '    sql &= " And Razao.Conta_Id = '" & Left(ddlGrupo.SelectedValue, 7) & "'" & vbCrLf
            '    Cliente = Right(hdfCliente.Value, 14)
            '    If Right(Cliente, 3) = "CPF" Then
            '        Cliente = Left(Cliente, 11)
            '    End If
            '    sql &= " And Razao.Cliente_Id = '" & Cliente & "'" & vbCrLf
            'End If

            If Not String.IsNullOrWhiteSpace(hdfCliente.Value) Then
                sql &= "And Razao.Cliente_Id = '" & hdfCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                    "   And Razao.EndCliente_Id = " & hdfCliente.Value.Split("-")(1) & vbCrLf
            End If

            If ddlGrupo.SelectedValue <> "" And ddlProduto.SelectedValue = "" Then
                sql &= " And GruposDeEstoques.Grupo_Id = '" & ddlGrupo.SelectedValue & "'" & vbCrLf
            End If

            If ddlProduto.SelectedValue <> "" Then
                sql &= " And Razao.Produto = '" & ddlProduto.SelectedValue & "'" & vbCrLf
            End If

            sql &= "           And (Razao.Movimento_Id < '" & txtDataInicial.Text.ToSqlDate() & "')" & vbCrLf &
                   "  GROUP BY Razao.Empresa_Id, Razao.EndEmpresa_Id, Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, PlanoDeContas.Titulo, Empresa.Reduzido," & vbCrLf &
                   "                     Empresa.Nome , Empresa.Cidade" & vbCrLf &
                   "  Union" & vbCrLf &
                   "  SELECT     Empresa.Reduzido AS RedEmpresa, Razao.Empresa_Id AS Empresa, Razao.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa," & vbCrLf &
                   "                     Empresa.Cidade AS CidadeEmpresa, Razao.Conta_Id AS Conta, PlanoDeContas.Titulo AS NomeConta, Razao.Cliente_Id AS Cliente," & vbCrLf &
                   "                     Razao.EndCliente_Id AS EndCliente, Clientes.Nome AS NomeCliente, Clientes.Cidade AS CidadeCliente, Razao.Movimento_Id AS Data," & vbCrLf &
                   "                     Razao.Lote_Id AS Lote, Razao.Sequencia_Id AS Sequencia, Razao.Titulo as Titulo, Razao.Produto AS Produto, Razao.Historico, Razao.DebitoOficial AS Debitos, Razao.CreditoOficial AS Creditos," & vbCrLf &
                   "                     0 AS Saldo" & vbCrLf &
                   "  FROM         GruposDeEstoques INNER JOIN" & vbCrLf &
                   "                    Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo RIGHT OUTER JOIN" & vbCrLf &
                   "                    Razao ON Produtos.Produto_Id = Razao.Produto LEFT OUTER JOIN" & vbCrLf &
                   "                    Clientes AS Empresa ON Razao.Empresa_Id = Empresa.Cliente_Id AND Razao.EndEmpresa_Id = Empresa.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
                   "                    Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id AND Razao.EndCliente_Id = Clientes.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
                   "                    PlanoDeContas ON Razao.Conta_Id = PlanoDeContas.Conta_Id" & vbCrLf &
                   "                    WHERE 1=1 " & vbCrLf

            If chkConsolidarEmpresa.Checked Then
                sql &= " AND (left(Razao.Empresa_Id,8) = '" & Left(ddlEmpresa.SelectedValue, 8) & "') " & vbCrLf
            Else
                sql &= " AND (Razao.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "') " & vbCrLf
            End If

            If hdfConta.Value <> "" Then
                sql &= " And Razao.Conta_Id = '" & Left(hdfConta.Value, 9) & "'" & vbCrLf
            End If

            If ddlGrupo.SelectedValue <> "" And hdfConta.Value = "" And hdfCliente.Value = "" Then
                sql &= " And Razao.Conta_Id Like '" & Left(ddlGrupo.SelectedValue, 7) & "%'" & vbCrLf
            End If

            'If hdfCliente.Value <> "" Then
            '    sql &= " And Razao.Conta_Id = '" & Left(ddlGrupo.SelectedValue, 7) & "'" & vbCrLf
            '    Cliente = Right(hdfCliente.Value, 14)
            '    If Right(Cliente, 3) = "CPF" Then
            '        Cliente = Left(Cliente, 11)
            '    End If
            '    sql &= " And Razao.Cliente_Id = '" & Cliente & "'" & vbCrLf
            'End If

            If Not String.IsNullOrWhiteSpace(hdfCliente.Value) Then
                sql &= "And Razao.Cliente_Id = '" & hdfCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                    "   And Razao.EndCliente_Id = " & hdfCliente.Value.Split("-")(1) & vbCrLf
            End If

            If ddlGrupo.SelectedValue <> "" And ddlProduto.SelectedValue = "" Then
                sql &= " And GruposDeEstoques.Grupo_Id = '" & ddlGrupo.SelectedValue & "'" & vbCrLf
            End If

            If ddlProduto.SelectedValue <> "" Then
                sql &= " And Razao.Produto = '" & ddlProduto.SelectedValue & "'" & vbCrLf
            End If

            sql &= "    And (Razao.Movimento_Id  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
                   "  Order by Empresa, Data, Lote, Sequencia" & vbCrLf

            Return Banco.ConsultaDataSet(sql, "ExtratoDeConta")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Protected Sub ckPorCliente_CheckedChanged(sender As Object, e As EventArgs) Handles ckPorCliente.CheckedChanged
        Try
            chkHistorico.Checked = False
            chkContasDoPeriodo.Checked = False

            If ckPorCliente.Checked Then
                lnkPdf.Parent.Visible = False
            Else
                lnkPdf.Parent.Visible = True

            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkContasDoPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles chkContasDoPeriodo.CheckedChanged
        Try
            ckPorCliente.Checked = False
            chkHistorico.Checked = False
            hdfConta.Value = String.Empty
            txtConta.Text = String.Empty
            hdfCliente.Value = String.Empty
            txtCliente.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class