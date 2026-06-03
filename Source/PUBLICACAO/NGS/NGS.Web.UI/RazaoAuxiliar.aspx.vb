Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RazaoAuxiliar
    Inherits BasePage

    Private Sql As String
    Private i As Integer = 0
    Private DS As DataSet
    Private VCCustos As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Contabil)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RazaoAuxiliar", "ACESSAR") Then
                    Limpar()
                    CargaCCusto()
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    ucConsultaProduto.SetarHID(HID.Value)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = Date.Now.ToString("01/MM/yyyy")
        txtDataFinal.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataEmissao.Text = Date.Now.ToString("dd/MM/yyyy")
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "")
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
        CargaCCusto()
        txtNomeCliente.Text = ""
        CodigoCliente.Value = ""
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Function Validar()
        If Not Len(txtGrupoInicial.Text) = Len(txtGrupoFinal.Text) Then
            MsgBox(Me.Page, "Digito da conta inicial e final devem conter o mesmo tamanho.")
            Return False
        End If
        Return True
    End Function

    Protected Sub ImgContaInicial_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgContaInicial.Click
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objContaInicial" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgContaFinal_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles ImgContaFinal.Click
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objContaFinal" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objContaInicial" & HID.Value) IsNot Nothing Then
            txtGrupoInicial.Text = CType(HttpContext.Current.Session("objContaInicial" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
            Session.Remove("objContaInicial" & HID.Value)
        ElseIf Session("objContaFinal" & HID.Value) IsNot Nothing Then
            txtGrupoFinal.Text = CType(HttpContext.Current.Session("objContaFinal" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
            Session.Remove("objContaFinal" & HID.Value)
        ElseIf Session("objClienteRazao" & HID.Value) IsNot Nothing Then
            Dim cli As [Lib].Negocio.Cliente = CType(Session("objClienteRazao" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(cli)
            With cli
                txtNomeCliente.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                CodigoCliente.Value = .Codigo & " - " & .CodigoEndereco
            End With
            Session.Remove("objClienteRazao" & HID.Value)
            'ElseIf Session("objProdutoRazao" & HID.Value) IsNot Nothing Then
            '    Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoRazao" & HID.Value)
            '    txtProduto.Text = objProduto.Nome
            '    CodigoProduto.Value = objProduto.Codigo
            '    Session.Remove("objProdutoRazao" & HID.Value)
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRazao" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoRazao" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub ConcatenaDataSetSaldoAnterior(ByVal dsRazaoAuxiliar As DataSet)
        Dim drNew As DataRow
        Dim dc As New DataColumn("SaldoAnterior", GetType(Double))
        Dim strConta As String = ""
        Dim strCliente As String = ""

        dc.AllowDBNull = False
        dc.DefaultValue = 0
        dsRazaoAuxiliar.Tables("Razao").Columns.Add(dc)

        For intPos As Integer = 0 To dsRazaoAuxiliar.Tables(0).Rows.Count - 1
            Dim drRegistros As DataRow = dsRazaoAuxiliar.Tables(0).Rows(intPos)
            If strConta <> drRegistros("Conta_Id") And strCliente = drRegistros("Cliente_Id") & Format(drRegistros("EndCliente_Id"), "000") Then
                strConta = drRegistros("Conta_Id")

                If UCase(drRegistros("Historico")) <> "SALDO ANTERIOR" Then
                    drNew = dsRazaoAuxiliar.Tables(0).NewRow

                    drNew("Conta_Id") = drRegistros("Conta_Id")
                    drNew("NomeConta") = drRegistros("NomeConta")
                    drNew("Movimento_Id") = "01/01/01"
                    drNew("Lote_Id") = 0
                    drNew("Sequencia_Id") = 0
                    drNew("Titulo") = 0
                    drNew("Historico") = "Saldo Anterior"
                    drNew("Debito") = 0
                    drNew("Credito") = 0
                    drNew("Cliente_Id") = drRegistros("Cliente_Id")
                    drNew("EndCliente_Id") = drRegistros("EndCliente_Id")
                    drNew("NomeCliente") = drRegistros("NomeCliente")

                    dsRazaoAuxiliar.Tables("Razao").Rows.Add(drNew)
                End If
            End If
        Next
    End Sub

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

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RazaoAuxiliar", "RELATORIO") Then
                If Validar() Then
                    Dim Ds_RazaoAuxiliarAux As New DataSet

                    Parametros()

                    Dim LotesDesconsiderados As String = VerIsolarLotes()
                    Dim UnidadeDeNegocio As String = ddlUnidadeDeNegocio.SelectedValue
                    Dim Empresa As String() = ddlEmpresa.SelectedValue.Split("-")
                    Dim GrupoInicial As String = txtGrupoInicial.Text
                    Dim GrupoFinal As String = txtGrupoFinal.Text
                    Dim Moeda As String = ddlMoeda.SelectedValue
                    Dim DataInicial As String = txtDataInicial.Text
                    Dim DataFinal As String = txtDataFinal.Text
                    'Dim CodCusto As String = ddlCentroDeCusto.SelectedValue

                    If Len(GrupoInicial) < 9 Then
                        GrupoInicial &= StrDup(9 - Len(GrupoInicial), "0")
                    End If
                    If Len(GrupoFinal) < 9 Then
                        GrupoFinal &= StrDup(9 - Len(GrupoFinal), "9")
                    End If

                    Dim strSQL As String
                    Dim ColunaBanco As String
                    If Moeda = "REAL" Then
                        ColunaBanco = "Oficial"
                    Else
                        ColunaBanco = "Moeda"
                    End If

                    strSQL = "SELECT R.conta_id + replicate('0',9 - len(R.conta_Id)) as Conta, R.Conta_Id, PC.Titulo AS NomeConta, R.Movimento_Id, " & vbCrLf

                    If chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked Then _
                        strSQL &= "R.Empresa_Id, R.EndEmpresa_Id, E.Nome as NomeEmp, E.Cidade as CidadeEmp, E.Estado as EstadoEmp,"


                    strSQL &= "      R.Lote_Id, R.Sequencia_Id, R.Titulo, R.Historico, " & vbCrLf & _
                             "       COALESCE(R.Debito" & ColunaBanco & ", 0) AS Debito, COALESCE(R.Credito" & ColunaBanco & ", 0) AS Credito, " & vbCrLf & _
                             "       ISNULL(R.Cliente_Id, '') AS Cliente_Id, ISNULL(R.EndCliente_Id, '') AS EndCliente_Id, " & vbCrLf & _
                             "       ISNULL(C.Nome, '') AS NomeCliente,R.Custo " & vbCrLf & _
                             " Into #RazTemp" & vbCrLf & _
                             "  FROM Razao R WITH (NOLOCK) " & vbCrLf & _
                             " INNER JOIN PlanoDeContas PC " & vbCrLf & _
                             "    ON R.Conta_Id = PC.Conta_Id " & vbCrLf

                    If chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked Then
                        strSQL &= " Inner JOIN Clientes E " & vbCrLf & _
                                 "    ON E.Cliente_Id  = R.Empresa_Id " & vbCrLf & _
                                 "   AND E.Endereco_Id = R.EndEmpresa_Id " & vbCrLf
                    End If

                    strSQL &= "  LEFT JOIN Clientes C " & vbCrLf & _
                             "    ON C.Cliente_Id = R.Cliente_Id " & vbCrLf & _
                             "   AND C.Endereco_Id = R.EndCliente_Id " & vbCrLf & _
                             "  LEFT JOIN Produtos Pr " & vbCrLf & _
                             "    ON R.Produto = Pr.Produto_Id " & vbCrLf & _
                             " WHERE len(R.Conta_Id) > 0" & vbCrLf & _
                             "   AND R.conta_id + replicate('0',9 - len(R.conta_Id)) >= '" & GrupoInicial & "'" & vbCrLf & _
                             "   AND R.conta_id + replicate('0',9 - len(R.conta_Id)) <= '" & GrupoFinal & "'" & vbCrLf & _
                             "   AND R.Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "' " & vbCrLf

                    If LotesDesconsiderados.Length > 0 Then
                        strSQL &= "   AND R.LOTE_ID NOT IN " & LotesDesconsiderados & vbCrLf
                    End If

                    If UnidadeDeNegocio.Length > 0 Then strSQL &= " AND isnull(R.UnidadeDeNegocio,'" & UnidadeDeNegocio & "') = '" & UnidadeDeNegocio & "'"
                    If Empresa.Length > 1 Then
                        If chkConsolidarEmpresa.Checked Then
                            strSQL &= " AND left(R.Empresa_Id,8) = '" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
                        Else
                            strSQL &= " AND R.Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                                      " AND R.EndEmpresa_Id = " & Empresa(1) & vbCrLf
                        End If
                    End If
                    If txtNomeCliente.Text.Length > 0 Then
                        Dim cliente As String() = CodigoCliente.Value.Split("-")
                        strSQL &= " AND R.Cliente_Id " & IIf(chkCnpjCpf.Checked, "LIKE '" & Mid(cliente(0), 1, 8) & "%'", " = '" & cliente(0) & "' and R.EndCliente_Id = " & cliente(1)) & vbCrLf
                    End If
                    'If CodCusto.Length > 0 Then strSQL &= "AND R.Custo = '" & CodCusto & "' " & vbCrLf
                    If VCCustos <> "" Then strSQL &= " And R.Custo In " & VCCustos

                    'If txtProduto.Text.Length > 0 Then strSQL &= "AND R.Produto = '" & CodigoProduto.Value & "' " & vbCrLf

                    Dim RetornoProdutos As ArrayList

                    If ucSelecaoProduto.TemSelecionado Then
                        RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Pr.Grupo", "Pr.Produto_Id", "")
                        strSQL &= " AND " & RetornoProdutos(0)
                        'ParametrosRelatorio &= RetornoProdutos(1)
                    End If


                    If txtLote.Text.Trim.Length > 0 Then strSQL &= "AND R.Lote_Id = " & txtLote.Text & vbCrLf

                    strSQL &= "Insert Into #RazTemp " & vbCrLf & _
                              "SELECT R.conta_id + replicate('0',9 - len(R.conta_Id)) as Conta," & vbCrLf & _
                              "       R.Conta_Id," & vbCrLf & _
                              "       PC.Titulo AS NomeConta, " & vbCrLf & _
                              "       convert(datetime,'" & DataInicial.ToSqlDate() & "') - 1 AS Movimento_Id," & vbCrLf

                    If chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked Then _
                           strSQL &= "       R.Empresa_Id, R.EndEmpresa_Id, E.Nome as NomeEmp, E.Cidade as CidadeEmp, E.Estado as EstadoEmp," & vbCrLf

                    strSQL &= "       0 AS Lote_Id," & vbCrLf & _
                             "       0 AS Sequencia_Id," & vbCrLf & _
                             "       0 AS Titulo," & vbCrLf & _
                             "      'SALDO ANTERIOR' AS Historico, " & vbCrLf & _
                             "      SUM(R.Debito" & ColunaBanco & ") AS Debito," & vbCrLf & _
                             "      SUM(Credito" & ColunaBanco & ") AS Credito, " & vbCrLf & _
                             "      ISNULL(R.Cliente_Id, '') AS Cliente_Id," & vbCrLf & _
                             "      ISNULL(R.EndCliente_Id, '') AS EndCliente_Id, " & vbCrLf & _
                             "      ISNULL(C.Nome, '') AS NomeCliente,0 as Custo " & vbCrLf & _
                             "  FROM Razao R WITH (NOLOCK)" & vbCrLf & _
                             " INNER JOIN PlanoDeContas PC " & vbCrLf & _
                             "    ON R.Conta_Id = PC.Conta_Id " & vbCrLf

                    If chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked Then
                        strSQL &= " Inner JOIN Clientes E " & vbCrLf & _
                                  "    ON E.Cliente_Id  = R.Empresa_Id " & vbCrLf & _
                                  "   AND E.Endereco_Id = R.EndEmpresa_Id " & vbCrLf
                    End If

                    strSQL &= "  LEFT JOIN Clientes C " & vbCrLf & _
                             "    ON C.Cliente_Id  = R.Cliente_Id " & vbCrLf & _
                             "   AND C.Endereco_Id = R.EndCliente_Id " & vbCrLf & _
                             " WHERE len(R.Conta_Id) > 0" & vbCrLf & _
                             "   AND R.conta_id + replicate('0',9 - len(R.conta_Id)) >= '" & GrupoInicial & "'" & vbCrLf & _
                             "   AND R.conta_id + replicate('0',9 - len(R.conta_Id)) <= '" & GrupoFinal & "'" & vbCrLf & _
                             "   AND R.Movimento_Id < '" & DataInicial.ToSqlDate() & "' " & vbCrLf

                    If LotesDesconsiderados.Length > 0 Then strSQL &= "   AND R.LOTE_ID NOT IN " & LotesDesconsiderados & vbCrLf

                    If UnidadeDeNegocio.Length > 0 Then strSQL &= " AND isnull(R.UnidadeDeNegocio,'" & UnidadeDeNegocio & "') = '" & UnidadeDeNegocio & "'" & vbCrLf

                    If Empresa.Length > 1 Then
                        If chkConsolidarEmpresa.Checked Then
                            strSQL &= " AND left(R.Empresa_Id,8) = '" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
                        Else
                            strSQL &= " AND R.Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                                      " AND R.EndEmpresa_Id = " & Empresa(1) & vbCrLf
                        End If
                    End If

                    If txtNomeCliente.Text.Length > 0 Then
                        Dim cliente As String() = CodigoCliente.Value.Split("-")
                        strSQL &= " AND R.Cliente_Id " & IIf(chkCnpjCpf.Checked, "LIKE '" & Mid(cliente(0), 1, 8) & "%'", " = '" & cliente(0) & "' and R.EndCliente_Id = " & cliente(1)) & vbCrLf
                    End If

                    'If CodCusto.Length > 0 Then strSQL &= "AND R.Custo = '" & CodCusto & "' " & vbCrLf
                    If VCCustos <> "" Then strSQL &= " And R.Custo In " & VCCustos
                    'If txtProduto.Text.Length > 0 Then strSQL &= "AND R.Produto = '" & CodigoProduto.Value & "' " & vbCrLf
                    If txtLote.Text.Trim.Length > 0 Then strSQL &= " AND Lote_Id = " & txtLote.Text & vbCrLf

                    strSQL &= " GROUP BY R.Conta_Id, PC.Titulo, R.Cliente_Id, R.EndCliente_Id, C.Nome" & IIf(chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked, ", R.Empresa_Id, R.EndEmpresa_Id, E.Nome, E.Cidade, E.Estado", "") & vbCrLf

                    strSQL &= " Select Conta_Id, NomeConta, Movimento_Id,"

                    If chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked Then
                        strSQL &= " Empresa_Id, EndEmpresa_Id, NomeEmp, CidadeEmp, EstadoEmp,"
                    End If

                    strSQL &= " Lote_Id, Sequencia_Id, Titulo, Historico, " & vbCrLf & _
                              " Debito, Credito, Cliente_Id, EndCliente_Id, NomeCliente, Custo " & vbCrLf & _
                              " From #RazTemp " & vbCrLf

                    If chkContasComMovimento.Checked Then
                        strSQL &= " Where  Conta_id + cliente_id+convert(nvarchar,endcliente_Id)  in (select conta_id+cliente_id+convert(nvarchar,endcliente_Id) " & vbCrLf & _
                                                                                                    " from #RazTemp " & vbCrLf & _
                                                                                                    " where historico <> 'SALDO ANTERIOR' " & vbCrLf & _
                                                                                                    " group by conta_id,cliente_id,endcliente_id " & vbCrLf & _
                                                                                                    " having count(*)>0) " & vbCrLf
                    End If

                    strSQL &= " ORDER BY Conta_Id, Movimento_Id, " & IIf(chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked, "Empresa_Id, EndEmpresa_Id, ", "") & "Lote_Id, Cliente_Id, EndCliente_Id"

                    Ds_RazaoAuxiliarAux = Banco.ConsultaDataSet(strSQL, "Razao")
                    ConcatenaDataSetSaldoAnterior(Ds_RazaoAuxiliarAux)

                    Dim namereport As String = "Cr_RazaoAuxiliar"

                    If chkConsolidarEmpresa.Checked AndAlso rbConsolidarEmpresaAnalitico.Checked Then namereport = "Cr_RazaoAuxiliarConsolidado"

                    Dim crpt As New ReportDocument()
                    Try
                        crpt.FileName = Server.MapPath("~/Reports/" & namereport & ".rpt")
                        crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)
                        Dim NomeArquivo2 As String
                        If Pdf = True Then
                            NomeArquivo2 = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                        Else
                            NomeArquivo2 = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
                        End If
                        Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                        Dim arquivo As String = NomeArquivo

                        crpt.SetDataSource(Ds_RazaoAuxiliarAux)

                        Dim objEmpresa As [Lib].Negocio.Cliente = Nothing
                        If Empresa.Length > 1 Then objEmpresa = New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))

                        Dim crparametervalues As ParameterValues
                        Dim crparameterdiscretevalue As ParameterDiscreteValue
                        Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                        Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                        crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Empresa")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        If Empresa.Length > 1 Then
                            crparameterdiscretevalue.Value = objEmpresa.Codigo
                        Else
                            crparameterdiscretevalue.Value = ""
                        End If
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        If Empresa.Length > 1 Then
                            crparameterdiscretevalue.Value = objEmpresa.Nome
                        Else
                            crparameterdiscretevalue.Value = "Todas as Empresas"
                        End If
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        If Empresa.Length > 1 Then
                            crparameterdiscretevalue.Value = objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                        Else
                            crparameterdiscretevalue.Value = ""
                        End If
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        If Moeda = "Dolar" Then
                            crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
                            crparametervalues = crparameterfielddefinition.CurrentValues
                            crparameterdiscretevalue = New ParameterDiscreteValue
                            crparameterdiscretevalue.Value = "Razão Auxiliar em Dolar"
                            crparametervalues.Add(crparameterdiscretevalue)
                            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                        Else
                            crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
                            crparametervalues = crparameterfielddefinition.CurrentValues
                            crparameterdiscretevalue = New ParameterDiscreteValue
                            crparameterdiscretevalue.Value = "Razão Auxiliar em Real"
                            crparametervalues.Add(crparameterdiscretevalue)
                            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)
                        End If

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Registro")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Título"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("UNNegocio")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = IIf(ddlUnidadeDeNegocio.SelectedIndex = 0, "Todas", ddlUnidadeDeNegocio.SelectedItem.Text)
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Moeda")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "tipo"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Periodo")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Período: " & DataInicial & " " & DataFinal
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("RotEmissao")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Emissão: "
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Emissao")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "01/01/2006"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("startPage")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = IIf(txtInicialFolha.Text.Length = 0, 0, txtInicialFolha.Text)
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Pagina")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Folha.: "
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Data")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Data"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Lote")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Lote"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Sequencia")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Seq"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Historico")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Histórico"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Debito")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Débito"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Credito")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Crédito"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Saldo")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Saldo"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("SubSaldo")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "Saldo Final"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        crparameterfielddefinition = crparameterfielddefinitions.Item("Custo")
                        crparametervalues = crparameterfielddefinition.CurrentValues
                        crparameterdiscretevalue = New ParameterDiscreteValue
                        crparameterdiscretevalue.Value = "C.Custo"
                        crparametervalues.Add(crparameterdiscretevalue)
                        crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                        If Pdf = True Then
                            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
                        Else
                            Dim CrExportOptions As New ExportOptions()
                            Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                            Dim CrFormatTypeOptions As New ExcelFormatOptions()
                            crpt.ExportToDisk(ExportFormatType.Excel, arquivo)

                            With CrExportOptions
                                .ExportDestinationType = ExportDestinationType.DiskFile
                                .ExportFormatType = ExportFormatType.Excel
                                .DestinationOptions = CrDiskFileDestinationOptions
                                .FormatOptions = CrFormatTypeOptions
                            End With
                            crpt.ExportToDisk(ExportFormatType.Excel, arquivo)
                        End If

                        If IO.File.Exists(arquivo) Then
                            If Pdf = True Then
                                ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
                            Else
                                ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, Guid.NewGuid().ToString, "window.location = '" & NomeArquivo2 & "';", True)
                            End If
                        End If
                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                    Finally
                        crpt.Close()
                        crpt.Dispose()
                    End Try
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaCCusto()
        Dim Total As Integer = 0

        Sql = "SELECT CentroDeCusto_Id AS CODIGO, Descricao as Descricao " & _
              "FROM CentrosDeCustos where Ativo = 1 " & _
              "ORDER BY CentroDeCusto_Id"
        i = 0
        DS = Banco.ConsultaDataSet(Sql, "CentrosDeCustos")

        Total = DS.Tables(0).Rows.Count

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            If i <= Total / 2 Then
                SelecaoDeGrupos1.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            Else
                SelecaoDeGrupos2.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            End If
            i += 1
        Next
    End Sub

    Private Sub VerSelecaoCCustos()
        VCCustos = "('"
        Dim i As Integer = 0
        For i = 0 To SelecaoDeGrupos1.Items.Count - 1
            If SelecaoDeGrupos1.Items(i).Selected Then
                If Len(VCCustos) > 2 Then
                    VCCustos &= ",'" & SelecaoDeGrupos1.Items(i).Value() & "'"
                Else
                    VCCustos &= SelecaoDeGrupos1.Items(i).Value() & "'"
                End If
            End If
        Next
        i = 0
        For i = 0 To SelecaoDeGrupos2.Items.Count - 1
            If SelecaoDeGrupos2.Items(i).Selected Then
                If Len(VCCustos) > 2 Then
                    VCCustos &= ",'" & SelecaoDeGrupos2.Items(i).Value() & "'"
                Else
                    VCCustos &= SelecaoDeGrupos2.Items(i).Value() & "'"
                End If
            End If
        Next
        If Len(VCCustos) > 2 Then
            VCCustos &= ")"
        Else
            VCCustos = ""
        End If
    End Sub

    Private Sub Parametros()
        VerSelecaoCCustos()
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function VerIsolarLotes() As String
        Dim VIsolarLotes As String = "("
        Dim i As Integer
        For i = 0 To IsolarLotes.Items.Count - 1
            If IsolarLotes.Items(i).Selected Then
                If Len(VIsolarLotes) > 1 Then
                    VIsolarLotes &= "," & IsolarLotes.Items(i).Value()
                Else
                    VIsolarLotes &= IsolarLotes.Items(i).Value()
                End If
            End If
        Next
        If Len(VIsolarLotes) > 1 Then
            Return VIsolarLotes & ")"
        Else
            Return ""
        End If
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RazaoAuxiliar")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class