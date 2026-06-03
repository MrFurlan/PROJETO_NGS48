Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ConciliacaoDeContas
    Inherits BasePage

    Dim Sql As String
    Dim Cliente As String
    Dim campo() As String
    Dim linha As String
    Dim cl As Integer = 0
    Dim Cg As Integer
    Dim Mensagem As String
    Dim Opcao As String
    Dim pagina As Integer = 0
    Dim Aquisicao As Date
    Dim Uso As Date
    Dim valor As Double
    Dim ds As DataSet
    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim PaginaInicial As Integer = 0
    Dim ClienteNome As String
    Dim ClienteInscricao As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ConciliacaoDeContas", "ACESSAR") Then
                BuncarUnidadeDeNegocio
                CargaCentroDeCusto()
                Limpar()
                txtDataInicial.Text = Format(Today, "01/01/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            hdfConta.Value = ""
            txtConta.Text = ""
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Private Sub CargaUnidade()
    '    Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf & _
    '          "FROM Clientes C " & vbCrLf & _
    '          "INNER JOIN ClientesXTipos CT " & vbCrLf & _
    '          "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf & _
    '          "WHERE CT.Tipo_Id = 050 " & vbCrLf & _
    '          "ORDER BY Nome" & vbCrLf
    '    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
    '        ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
    '    Next

    '    ddlUnidade.Items.Insert(0, "")
    '    ddlUnidade.SelectedIndex = 0
    'End Sub

    'Private Sub VerificaUnidade()
    '    Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
    '          "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
    '          "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
    '          " from Usuarios" & vbCrLf & _
    '          " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf
    '    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
    '        ddlUnidade.SelectedValue = Dr("AcessoUnidade")
    '        CargaEmpresas()
    '        ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
    '    Next
    'End Sub

    'Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        CargaEmpresas()
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Private Sub CargaEmpresas()
    '    Dim Codigo As String
    '    Dim Descricao As String
    '    Dim Nome As String
    '    Dim Cidade As String
    '    Dim Cnpj As String

    '    ddlEmpresa.Items.Clear()

    '    HttpContext.Current.Session("EmpresaIcms") = ""
    '    HttpContext.Current.Session("ProcessoIcms") = 0

    '    Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
    '          " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
    '          " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
    '          " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf
    '    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
    '        Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

    '        Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
    '        Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

    '        Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
    '        Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
    '        Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

    '        ddlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
    '    Next
    '    ddlEmpresa.Items.Insert(0, "")
    '    ddlEmpresa.SelectedIndex = 0
    'End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objCliConciliacaoDeContas" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliConciliacaoDeContas" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliConciliacaoDeContas" & HID.Value)
        ElseIf Not Session("objExtratoDeConta" & HID.Value) Is Nothing Then
            Dim objConta As [Lib].Negocio.PlanoDeConta = CType(obj, [Lib].Negocio.PlanoDeConta)
            hdfConta.Value = objConta.Conta
            txtConta.Text = objConta.Conta & "-" & objConta.Titulo
            Session.Remove("objExtratoDeConta" & HID.Value)
        End If
    End Sub

    Private Sub CargaCentroDeCusto()
        Sql = " SELECT CentroDeCusto_Id as Codigo, Descricao " & vbCrLf & _
                " FROM CentrosDeCustos" & vbCrLf & _
                "  where Ativo = 1 and Len(CentroDeCusto_Id) = 5" & vbCrLf & _
                " Order By CentroDeCusto_Id" & vbCrLf
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlCentroDeCusto.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
        DdlCentroDeCusto.Items.Insert(0, "")
        DdlCentroDeCusto.SelectedIndex = 0
    End Sub

    Private Function Validar()
        If ddlUnidade.Text = "" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf txtConta.Text = "" Then
            MsgBox(Me.Page, "Conta é obrigatório.")
            Return False
        ElseIf HttpContext.Current.Session("TemCliente") = "s" AndAlso txtCliente.Text = "" Then
            MsgBox(Me.Page, "Cliente é obrigatório.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Limpar()
        hdfConta.Value = ""
        txtConta.Text = ""

        txtCodigoCliente.Value = ""
        txtCliente.Text = ""

        DdlCentroDeCusto.SelectedIndex = 0

        ddlGrupo.SelectedIndex = 0
        ddlProduto.Items.Clear()

        lnkConsultar.Parent.Visible = True
        lnkConfirmar.Parent.Visible = False

        HttpContext.Current.Session("Titulo") = ""
        HttpContext.Current.Session("TemCliente") = ""
        HttpContext.Current.Session("TipoDeCliente") = ""
        HttpContext.Current.Session("TemProduto") = ""
        HttpContext.Current.Session("TemCusto") = ""
        HttpContext.Current.Session("Cliente") = ""
        HttpContext.Current.Session("ClienteEnd") = 0
        HttpContext.Current.Session("ClienteNome") = ""
        HttpContext.Current.Session("ClienteCidade") = ""
        HttpContext.Current.Session("ClienteEstado") = ""
        HttpContext.Current.Session("Reduzido") = ""

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaPlanoDeContas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

        grdConciliacao.DataSource = Nothing
        grdConciliacao.DataBind()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros: " & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Unidade: " & ddlUnidade.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtConta.Text) Then
            param &= "Conta: " & txtConta.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            param &= "Cliente: " & txtCliente.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à: " & txtDataFinal.Text
        End If

        Return param
    End Function

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("ConciliacaoDeContas", "LEITURA") Then
            Validar()

            If Mensagem = "" Then
                Dim strSQL As String
                Dim ds As New DataSet

                Cliente = ddlEmpresa.SelectedValue
                campo = Cliente.Split("-")

                Dim Empresa As String = campo(0)
                Dim EndEmpresa As String = campo(1)

                Dim Conta As String = hdfConta.Value
                Cliente = HttpContext.Current.Session("Cliente")
                Dim EndCliente As String = HttpContext.Current.Session("ClienteeND")

                Dim Custo As String = ""
                Dim Produto As String = ""
                Dim Aberto As String = "S"
                Dim Confirmado As String = "S"
                If rdAberto.Checked = False Then
                    Aberto = "N"
                End If
                If rdBaixados.Checked = False Then
                    Confirmado = "N"
                End If
                If Aberto = "N" And Confirmado = "N" Then
                    Confirmado = "S"
                    Aberto = "S"
                End If

                If EndCliente = "" Then
                    EndCliente = 0
                End If

                If Custo = "" Then
                    Custo = 0
                End If

                If Produto = "" Then
                    Produto = 0
                End If

                Dim DataInicial As String = Format(CDate(txtDataInicial.Text), "yyyy/MM/dd")
                Dim DataFinal As String = Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")

                strSQL = " SELECT  SUM(DebitoOficial) AS Debito, SUM(CreditoOficial) AS Credito" & vbCrLf & _
                         " Into #Anterior" & vbCrLf & _
                         " FROM Razao" & vbCrLf & _
                         " WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & " And Conta_Id = '" & Conta & "' And " & vbCrLf & _
                         " Cliente_Id = '" & Cliente & "' And EndCliente_Id = " & EndCliente & vbCrLf
                If Custo <> "0" Then
                    strSQL &= " And Custo = " & Custo
                End If
                If Produto <> "0" Then
                    strSQL &= " And Produto = '" & Produto & "' "
                End If

                strSQL &= " And Movimento_Id < '" & DataInicial & "'"

                If Aberto = "S" And Confirmado = "N" Then
                    strSQL &= " And (Conciliacao <> 'B' or Conciliacao is null) "
                ElseIf Aberto = "N" And Confirmado = "S" Then
                    strSQL &= " And Conciliacao = 'B' "
                End If

                strSQL &= " SELECT  SUM(DebitoOficial) AS Debito, SUM(CreditoOficial) AS Credito"
                'strSQL &= " SELECT  0 AS Debito, 0 AS Credito"
                strSQL &= " Into #Posterior" & vbCrLf & _
                          " FROM Razao" & vbCrLf & _
                          " WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & " And Conta_Id = '" & Conta & "' And " & vbCrLf & _
                          " Cliente_Id = '" & Cliente & "' And EndCliente_Id = " & EndCliente & vbCrLf
                If Custo <> "0" Then
                    strSQL &= " And Custo = " & Custo
                End If
                If Produto <> "0" Then
                    strSQL &= " And Produto = '" & Produto & "' "
                End If

                strSQL &= " And Movimento_Id <= '" & DataFinal & "'"

                If Aberto = "S" And Confirmado = "N" Then
                    strSQL &= " And (Conciliacao <> 'B' or Conciliacao is null) "
                ElseIf Aberto = "N" And Confirmado = "S" Then
                    strSQL &= " And Conciliacao = 'B' "
                End If

                '******************** saldo confirmado inicio
                strSQL &= " SELECT  SUM(DebitoOficial) AS Debito, SUM(CreditoOficial) AS Credito" & vbCrLf & _
                          " Into #AnteriorConfirmado" & vbCrLf & _
                          " FROM Razao" & vbCrLf & _
                          " WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & " And Conta_Id = '" & Conta & "' And " & vbCrLf & _
                          " Cliente_Id = '" & Cliente & "' And EndCliente_Id = " & EndCliente & vbCrLf
                If Custo <> "0" Then
                    strSQL &= " And Custo = " & Custo
                End If
                If Produto <> "0" Then
                    strSQL &= " And Produto = '" & Produto & "' "
                End If

                strSQL &= " And Movimento_Id < '" & DataInicial & "' And Conciliacao = 'B'"

                strSQL &= " SELECT  SUM(DebitoOficial) AS Debito, SUM(CreditoOficial) AS Credito"
                'strSQL &= " SELECT  0 AS Debito, 0 AS Credito"
                strSQL &= " Into #PosteriorConfirmado" & vbCrLf & _
                          " FROM Razao" & vbCrLf & _
                          " WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & " And Conta_Id = '" & Conta & "' And " & vbCrLf & _
                          " Cliente_Id = '" & Cliente & "' And EndCliente_Id = " & EndCliente & vbCrLf
                If Custo <> "0" Then
                    strSQL &= " And Custo = " & Custo
                End If
                If Produto <> "0" Then
                    strSQL &= " And Produto = '" & Produto & "' "
                End If

                strSQL &= " And Movimento_Id <= '" & DataFinal & "' And Conciliacao = 'B'"

                '******************** saldo confirmado fim.
                If Confirmado = "S" Then
                    strSQL &= "SELECT 00 AS STATUS, NULL as Movimento_Id, ' ' as Lote_Id, ' ' as Sequencia_Id, ' ' as Titulo, ' ' as Produto, ' ' as Consciliacao, NULL as DataDaBaixa , '0' as Custo, 'SALDO INICIAL CONFIRMADO' as Historico, isnull(Debito, 0) as Debito, isnull(Credito, 0) as Credito, 'C' AS Saldo, ' ' as Cliente_id , 0 as EndCliente_id " & vbCrLf & _
                              " FROM #AnteriorConfirmado " & vbCrLf

                    strSQL &= " Union "
                End If

                If Aberto = "S" Then
                    strSQL &= "SELECT 01 AS STATUS, NULL as Movimento_Id, ' ' as Lote_Id, ' ' as Sequencia_Id, ' ' as Titulo, ' ' as Produto, ' ' as Consciliacao, NULL as DataDaBaixa , '0' as Custo, 'SALDO INICIAL...........' as Historico, isnull(Debito, 0) as Debito, isnull(Credito, 0) as Credito, 'A' AS Saldo, ' ' as Cliente_id , 0 as EndCliente_id  " & vbCrLf & _
                              " FROM #Anterior " & vbCrLf

                    strSQL &= " Union "
                End If

                strSQL &= " SELECT  02 AS STATUS, Movimento_Id, REPLICATE('0', 4 - LEN(CAST(Lote_Id AS varchar))) + CAST(Lote_Id AS varchar) as Lote_ID, " & vbCrLf & _
                          " Sequencia_Id, Titulo, isnull(Produto, '') as Produto, " & vbCrLf & _
                          " (CASE  WHEN Conciliacao='B' THEN 'B'  ELSE 'A' end) as Conciliacao, DataDaBaixa,  " & vbCrLf & _
                          " Custo, Historico, DebitoOficial as Debito, CreditoOficial as Credito, 'D' As Saldo, Cliente_id , EndCliente_id " & vbCrLf

                strSQL &= " FROM Razao"

                strSQL &= " WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & " And Conta_Id = '" & Conta & "' And "
                strSQL &= " Cliente_Id = '" & Cliente & "' And EndCliente_Id = " & EndCliente
                If Custo <> "0" Then
                    strSQL &= " And Custo = " & Custo
                End If
                If Produto <> "0" Then
                    strSQL &= " And Produto = '" & Produto & "' "
                End If
                strSQL &= " And Movimento_Id BETWEEN '" & DataInicial & "' AND '" & DataFinal & "'" & vbCrLf & _
                          " " & vbCrLf
                If Aberto = "S" And Confirmado = "N" Then
                    strSQL &= " And (Conciliacao <> 'B' or Conciliacao is null) "
                ElseIf Aberto = "N" And Confirmado = "S" Then
                    strSQL &= " And Conciliacao = 'B' "
                End If
                If Aberto = "S" Then

                    strSQL &= " Union "

                    strSQL &= "SELECT 03 AS STATUS, NULL as Movimento_Id, ' ' as Lote_Id, ' ' as Sequencia_Id, ' ' as Titulo, ' ' as Produto, ' ' as Consciliacao, NULL as DataDaBaixa , '0' as Custo, 'SALDO FINAL.............' as Historico, isnull(Debito, 0) as Debito, isnull(Credito, 0) as Credito, 'B' AS Saldo, ' ' as Cliente_id , 0 as EndCliente_id  " & vbCrLf & _
                              " FROM #Posterior " & vbCrLf
                End If
                If Confirmado = "S" Then

                    strSQL &= " Union "

                    strSQL &= "SELECT 04 AS STATUS, NULL as Movimento_Id, ' ' as Lote_Id, ' ' as Sequencia_Id, ' ' as Titulo, ' ' as Produto, ' ' as Consciliacao, NULL as DataDaBaixa , '0' as Custo, 'SALDO FINAL CONFIRMADO..' as Historico, isnull(Debito, 0) as Debito, isnull(Credito, 0) as Credito, 'B' AS Saldo , ' ' as Cliente_id , 0 as EndCliente_id " & vbCrLf & _
                              " FROM #PosteriorConfirmado " & vbCrLf
                End If
                strSQL &= " ORDER BY 01, Movimento_Id, Lote_Id, Sequencia_Id "
                Try
                    ds = Banco.ConsultaDataSet(strSQL, "Razao")

                    If ds Is Nothing Then
                        MsgBox(Me.Page, "Período sem movimento...")
                    ElseIf ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Período sem movimento...")
                    Else
                        Dim dra As DataRow
                        Dim Saldo As Decimal = 0
                        Dim SaldoConfirmado As Decimal = 0
                        Dim debito1 As Decimal = 0
                        Dim debito3 As Decimal = 0
                        Dim credito1 As Decimal = 0
                        Dim credito3 As Decimal = 0
                        For Each dra In ds.Tables(0).Rows
                            If dra("Saldo").ToString() <> "B" Then Saldo += dra("Debito") - dra("Credito")
                            If dra("Saldo").ToString() = "C" Or dra("Consciliacao").ToString() = "B" Then SaldoConfirmado += dra("Debito") - dra("Credito")
                            If dra("Status").ToString() = "1" Then
                                debito1 = dra("debito")
                                credito1 = dra("credito")
                            End If
                            If dra("Status").ToString() = "3" Then
                                debito3 = dra("debito") - debito1
                                credito3 = dra("credito") - credito1
                                dra("debito") = debito3
                                dra("credito") = credito3
                            End If
                            If dra("Saldo") = "A" Then
                                'dra("Movimento_Id") = "NULL"
                                dra("Debito") = DBNull.Value
                                dra("Credito") = DBNull.Value
                            End If

                            Dim Valor As Decimal = 0

                            If dra("Status").ToString() = "4" Then Valor = SaldoConfirmado Else Valor = Saldo
                            Select Case Valor
                                Case Is > 0
                                    dra("Saldo") = Valor.ToString("N2") & "-D"
                                Case Is < 0
                                    dra("Saldo") = (Valor * -1).ToString("N2") & "-C"
                                Case Else
                                    dra("Saldo") = Valor.ToString("N2") & "DC"
                            End Select

                            If dra("Status") = "0" Then Saldo = 0
                        Next
                        grdConciliacao.DataSource = ds
                        grdConciliacao.DataBind()

                        lnkConsultar.Parent.Visible = False
                        lnkConfirmar.Parent.Visible = True

                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ConciliacaoDeContas", "RELATORIO") Then
                If Validar() Then
                    Dim strSql As String

                    Cliente = ddlEmpresa.SelectedValue
                    campo = Cliente.Split("-")

                    Dim Empresa As String = campo(0)
                    Dim EndEmpresa As String = campo(1)
                    Dim EmpresaNome As String = ""
                    Dim EmpresaCidade As String = ""
                    Dim EmpresaEstado As String = ""
                    Dim Conta As String = hdfConta.Value
                    Cliente = HttpContext.Current.Session("Cliente")
                    Dim EndCliente As String = HttpContext.Current.Session("ClienteEnd")
                    Dim ReduzidoCliente As String = HttpContext.Current.Session("Reduzido")
                    Dim NomeDoCliente As String = HttpContext.Current.Session("ClienteNome")

                    Dim Custo As String = ""
                    Dim Produto As String = ""
                    Dim DSRelatorio As New DataSet()

                    If EndCliente = "" Then
                        EndCliente = 0
                    End If

                    If Custo = "" Then
                        Custo = 0
                    End If

                    If Produto = "" Then
                        Produto = 0
                    End If
                    Dim Aberto As String = "S"
                    Dim Confirmado As String = "S"
                    If rdAberto.Checked = False Then
                        Aberto = "N"
                    End If
                    If rdBaixados.Checked = False Then
                        Confirmado = "N"
                    End If
                    If Aberto = "N" And Confirmado = "N" Then
                        Confirmado = "S"
                        Aberto = "S"
                    End If

                    Dim DataInicial As String = Format(CDate(txtDataInicial.Text), "yyyy/MM/dd")
                    Dim DataFinal As String = Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")
                    Dim PDataInicial As String = Format(CDate(txtDataInicial.Text), "dd/MM/yyyy")
                    Dim PDataFinal As String = Format(CDate(txtDataFinal.Text), "dd/MM/yyyy")
                    Dim TipoDeRelatorio As String = "PDF"

                    Const ABERTOS As String = "AND COALESCE(Conciliacao, 'A') <> 'B' "
                    Const BAIXADOS As String = "AND COALESCE(Conciliacao, 'A') = 'B' "

                    Dim strSQLSelect As String = "SELECT SUM(DebitoOficial) AS Debito, SUM(CreditoOficial) AS Credito "
                    Dim strSQLWhere As String = "FROM Razao " & _
                                                "WHERE Empresa_Id = '" & Empresa & "' " & _
                                                "AND EndEmpresa_Id = " & EndEmpresa & " " & _
                                                "AND Conta_Id = '" & Conta & "' "

                    If Cliente <> "" Then
                        strSQLWhere &= "AND Cliente_Id = '" & Cliente & "' " & _
                                       "AND EndCliente_Id = " & EndCliente & " "
                    End If

                    If Custo <> "0" Then strSQLWhere &= "AND Custo = " & Custo & " "
                    If Produto <> "0" Then strSQLWhere &= "AND Produto = '" & Produto & "' "

                    Dim strSQLSomaWhereConc As String = ""

                    If Aberto = "S" And Confirmado = "N" Then
                        strSQLSomaWhereConc &= ABERTOS
                    ElseIf Aberto = "N" And Confirmado = "S" Then
                        strSQLSomaWhereConc &= BAIXADOS
                    End If

                    strSql = strSQLSelect & vbCrLf & _
                             "INTO #Anterior " & vbCrLf & _
                             strSQLWhere & strSQLSomaWhereConc & vbCrLf & _
                             "AND Movimento_Id < '" & DataInicial & "' " & vbCrLf

                    strSql &= strSQLSelect & vbCrLf & _
                              "INTO #Posterior " & vbCrLf & _
                              strSQLWhere & strSQLSomaWhereConc & vbCrLf & _
                              "AND Movimento_Id <= '" & DataFinal & "' " & vbCrLf

                    strSql &= strSQLSelect & vbCrLf & _
                              "INTO #AnteriorConfirmado " & vbCrLf & _
                              strSQLWhere & BAIXADOS & vbCrLf & _
                              "AND Movimento_Id < '" & DataInicial & "' " & vbCrLf

                    strSql &= strSQLSelect & vbCrLf & _
                              "INTO #PosteriorConfirmado " & vbCrLf & _
                              strSQLWhere & BAIXADOS & vbCrLf & _
                              "AND Movimento_Id <= '" & DataFinal & "' " & vbCrLf

                    strSQLSelect = "SELECT NULL AS Movimento_Id, NULL AS Lote_Id, NULL AS Sequencia_Id, '' AS Produto, '0' AS CentroDeCustos, " & vbCrLf & _
                                  "'{HISTORICO}' AS Historico, COALESCE(Debito, 0) AS Debito, isnull(Credito, 0) AS Credito, " & vbCrLf & _
                                  "0.00 AS Saldo,  '' AS CD, NULL AS Titulo, NULL AS DataDaBaixa, '' AS Conciliacao, {STATUS} AS Status " & vbCrLf

                    If Confirmado = "S" Then
                        strSql &= strSQLSelect.Replace("{HISTORICO}", "SALDO INICIAL CONFIRMADO").Replace("{STATUS}", "0") & vbCrLf & _
                                  "FROM #AnteriorConfirmado " & vbCrLf & _
                                  "UNION " & vbCrLf
                    End If

                    If Aberto = "S" Then
                        strSql &= strSQLSelect.Replace("{HISTORICO}", "SALDO INICIAL").Replace("{STATUS}", "1") & vbCrLf & _
                                  "FROM #Anterior " & vbCrLf & _
                                  "UNION " & vbCrLf
                    End If

                    strSql &= "SELECT Movimento_Id, Lote_Id, Sequencia_Id, COALESCE(Produto, '') as Produto, " & vbCrLf & _
                              "Custo As CentroDeCustos, Historico, DebitoOficial AS Debito, CreditoOficial AS Credito, " & vbCrLf & _
                              "0.00 As Saldo, ' ' AS CD, Titulo as Titulo, " & vbCrLf & _
                              "DataDaBaixa, COALESCE(Conciliacao, 'A') AS Conciliacao, 2 AS Status " & vbCrLf & _
                              strSQLWhere & strSQLSomaWhereConc & vbCrLf
                    If rdBaixados.Checked Then
                        strSql &= "AND DataDaBaixa BETWEEN '" & DataInicial & "' AND '" & DataFinal & "' " & vbCrLf
                    Else
                        strSql &= "AND Movimento_Id BETWEEN '" & DataInicial & "' AND '" & DataFinal & "' " & vbCrLf
                    End If

                    If Aberto = "S" Then
                        strSql &= "UNION " & vbCrLf & _
                                  strSQLSelect.Replace("{HISTORICO}", "SALDO FINAL").Replace("{STATUS}", "3") & vbCrLf & _
                                  "FROM #Posterior " & vbCrLf
                    End If

                    If Confirmado = "S" Then
                        strSql &= "UNION " & vbCrLf & _
                        strSQLSelect.Replace("{HISTORICO}", "SALDO FINAL CONFIRMADO").Replace("{STATUS}", "4") & vbCrLf & _
                                  "FROM #PosteriorConfirmado " & vbCrLf
                    End If

                    strSql &= "ORDER BY 14, Movimento_Id, Lote_Id, Sequencia_Id "

                    ds = Banco.ConsultaDataSet(strSql, "Razao")

                    If ds Is Nothing Then
                        MsgBox(Me.Page, "Período sem movimento.", eTitulo.Info)
                        Return
                    ElseIf ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Período sem movimento.", eTitulo.Info)
                        Return
                    Else
                        Dim dra As DataRow
                        Dim Saldo As Decimal = 0
                        Dim SaldoConfirmado As Decimal = 0
                        Dim Debito1 As Decimal = 0
                        Dim Debito3 As Decimal = 0
                        Dim Credito1 As Decimal = 0
                        Dim Credito3 As Decimal = 0

                        For Each dra In ds.Tables(0).Rows
                            If Convert.ToInt32(dra("Status")) <> 3 Then Saldo += dra("Debito") - dra("Credito")
                            If Convert.ToInt32(dra("Status")) = 0 Or dra("Conciliacao").ToString() = "B" Then SaldoConfirmado += Convert.ToDouble(dra("Debito")) - Convert.ToDouble(dra("Credito"))
                            If Convert.ToInt32(dra("Status")) = 1 Then
                                Debito1 = dra("Debito")
                                Credito1 = dra("Credito")
                            End If
                            If dra("Status").ToString() = "3" Then
                                Debito3 = dra("debito") - Debito1
                                Credito3 = dra("credito") - Credito1
                                dra("debito") = Debito3
                                dra("credito") = Credito3
                            End If

                            If Convert.ToInt32(dra("Status")) = 1 Then
                                dra("Debito") = DBNull.Value
                                dra("Credito") = DBNull.Value
                            End If
                            'status = dra("Status").ToString()
                            'If Convert.ToInt32(dra("Status")) = 1 Then
                            '    Debito1 = dra("Debito")
                            '    Credito1 = dra("Credito")
                            'End If
                            'If dra("Status").ToString() = "3" Then
                            '    Debito3 = dra("debito") - Debito1
                            '    Credito3 = dra("credito") - Credito1
                            '    'dra("debito") = debito3
                            '    'dra("credito") = credito3
                            'End If


                            Dim Valor As Decimal = 0

                            If Convert.ToInt32(dra("Status")) = 4 Then Valor = SaldoConfirmado Else Valor = Saldo

                            Select Case Valor
                                Case Is > 0
                                    dra("Saldo") = Valor
                                    dra("CD") = "-D"
                                Case Is < 0
                                    dra("Saldo") = (Valor * -1)
                                    dra("CD") = "-C"
                                Case Else
                                    dra("Saldo") = Valor.ToString("N2")
                                    dra("CD") = ""
                            End Select

                            If Convert.ToInt32(dra("Status")) = 0 Then Saldo = 0
                        Next

                        DSRelatorio.Merge(ds)
                    End If

                    Dim strSQLBloq As String

                    strSQLBloq = "SELECT DataDaBaixa, SUM(ISNULL(DebitoOficial, 0)) AS Debito, " & vbCrLf & _
                                 "SUM(ISNULL(CreditoOficial, 0)) AS Credito " & vbCrLf & _
                                 "FROM Razao " & vbCrLf & _
                                 "LEFT JOIN Produtos " & vbCrLf & _
                                 "ON Produtos.Produto_Id = Razao.Produto " & vbCrLf & _
                    strSQLWhere.Replace("FROM Razao ", "") & vbCrLf & _
                                 "AND COALESCE(Conciliacao, 'A') <> 'B' " & vbCrLf & _
                                 "AND Movimento_Id Between '" & DataInicial & "'AND '" & DataFinal & "' " & vbCrLf & _
                                 "AND DataDaBaixa IS NOT NULL " & vbCrLf & _
                                 "GROUP BY DataDaBaixa, COALESCE(Conciliacao, 'A') " & vbCrLf

                    DSRelatorio.Merge(Banco.ConsultaDataSet(strSQLBloq, "Bloqueados"))

                    Dim strSQLComp As String

                    strSQLComp = "SELECT 0.00 AS Debito, SUM(ISNULL(CreditoOficial, 0)) AS Credito " & vbCrLf & _
                                 "FROM Razao " & vbCrLf & _
                                 "LEFT JOIN Produtos " & vbCrLf & _
                                 "ON Produtos.Produto_Id  = Razao.Produto " & vbCrLf & _
                    strSQLWhere.Replace("FROM Razao ", "") & vbCrLf & _
                                 "AND COALESCE(Conciliacao, 'A') <> 'B' " & vbCrLf & _
                                 "AND Movimento_Id <= '" & DataFinal & "' " & vbCrLf

                    DSRelatorio.Merge(Banco.ConsultaDataSet(strSQLComp, "Acompensar"))

                    Dim strSQLSoma As String

                    strSQLSoma = "SELECT SUM(DebitoOficial) AS Debito, SUM(CreditoOficial) AS Credito, " & vbCrLf & _
                                 "ISNULL(Conciliacao, 'A') AS Conciliacao " & vbCrLf & _
                    strSQLWhere & vbCrLf

                    If rdBaixados.Checked Then
                        strSQLSoma &= "AND DataDaBaixa BETWEEN '" & DataInicial & "' AND '" & DataFinal & "' " & vbCrLf
                    Else
                        strSQLSoma &= "AND Movimento_Id BETWEEN '" & DataInicial & "' AND '" & DataFinal & "' " & vbCrLf
                    End If

                    If Aberto = "S" And Confirmado = "N" Then
                        strSQLSoma &= "AND COALESCE(Conciliacao, 'A') <> 'B' "
                    ElseIf Aberto = "N" And Confirmado = "S" Then
                        strSQLSoma &= "AND COALESCE(Conciliacao, 'A') = 'B' "
                    End If

                    strSQLSoma &= "GROUP BY ISNULL(Conciliacao, 'A')"

                    DSRelatorio.Merge(Banco.ConsultaDataSet(strSQLSoma, "Soma"))

                    Sql = "  SELECT Nome, Cidade, Estado as Estado_id" & vbCrLf & _
                          " FROM   Clientes" & vbCrLf & _
                          " Where Cliente_Id = '" & Empresa & "' and Endereco_Id = " & EndEmpresa & vbCrLf

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Parametros", getParam())
                    parameters.Add("Titulo", "Relatório De Conciliação De Contas.")

                    Funcoes.BindReport(Me.Page, DSRelatorio, "Cr_Conciliacao", eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        If Funcoes.VerificaPermissao("ConciliacaoDeContas", "ALTERAR") Then

            Dim Cliente As String
            Dim Campo() As String
            Dim i As Integer
            Dim chkConciliacao As CheckBox
            Dim txtBaixa As TextBox
            Dim dtAux As DataTable
            Dim Cliente_id As HiddenField
            Dim EndCliente_id As HiddenField
            Dim SqlArray As New ArrayList

            dtAux = CType(HttpContext.Current.Session("CrCotacao"), DataTable)
            Cliente = ddlEmpresa.SelectedValue
            Campo = Cliente.Split("-")

            Dim Empresa As String = Campo(0)
            Dim EndEmpresa As String = Campo(1)

            Dim Conta As String = hdfConta.Value
            If Mensagem = "" Then
                i = 0
                While i < grdConciliacao.Rows.Count
                    chkConciliacao = CType(grdConciliacao.Rows(i).FindControl("chkConciliacao"), CheckBox)
                    txtBaixa = CType(grdConciliacao.Rows(i).FindControl("txtBaixa"), TextBox)
                    Cliente_id = CType(grdConciliacao.Rows(i).FindControl("Hiddenfield1"), HiddenField)
                    EndCliente_id = CType(grdConciliacao.Rows(i).FindControl("Hiddenfield2"), HiddenField)
                    Dim TxtBaixaFormatada As String

                    grdConciliacao.Rows(i).Cells(1).Text = RTrim(grdConciliacao.Rows(i).Cells(1).Text).Replace("&nbsp;", "")

                    If txtBaixa.Text.Trim().Length = 0 Then
                        TxtBaixaFormatada = "Null"
                    Else
                        Dim StrTxtBaixa As String() = txtBaixa.Text.Split("/")
                        TxtBaixaFormatada = "'" & StrTxtBaixa(2) & "-" & StrTxtBaixa(1) & "-" & StrTxtBaixa(0) & "'"
                    End If
                    If Not (chkConciliacao.Checked And txtBaixa.Text.Trim().Length = 0) Then
                        Sql = "UPDATE RAZAO "
                        Sql &= " set Conciliacao = '" & IIf(chkConciliacao.Checked, "B", "A") & "',"
                        Sql &= " DataDaBaixa = " & TxtBaixaFormatada & " "
                        Sql &= " WHERE Empresa_Id = '" & Empresa & "' And EndEmpresa_Id = " & EndEmpresa & " And Conta_Id = '" & Conta & "' And "
                        Sql &= " Cliente_Id = '" & Cliente_id.Value & "' And EndCliente_Id = " & EndCliente_id.Value
                        Sql &= " and Lote_Id = '" & grdConciliacao.Rows(i).Cells(2).Text & "' and Sequencia_Id = '" & grdConciliacao.Rows(i).Cells(3).Text & "'"
                        If grdConciliacao.Rows(i).Cells(1).Text.Length > 0 Then
                            Sql &= " and Movimento_Id = '" & grdConciliacao.Rows(i).Cells(1).Text.ToSqlDate() & "'"
                        End If
                        SqlArray.Add(Sql)
                    End If

                    i += 1
                End While

                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    grdConciliacao.DataSource = Nothing
                    grdConciliacao.DataBind()

                    lnkConsultar.Parent.Visible = True
                    lnkConfirmar.Parent.Visible = False
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para Alterar Registro")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConciliacaoDeContas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub btnConta_Click(sender As Object, e As EventArgs) Handles btnConta.Click
        HttpContext.Current.Session("ssCampo") = "ExtratoDeConta"
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "objExtratoDeConta" & HID.Value)
    End Sub

    Protected Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click
        If HttpContext.Current.Session("TemCliente" & HID.Value) = "S" Then
            HttpContext.Current.Session("ssCampo") = "ExtratoDeConta"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliConciliacaoDeContas" & HID.Value)
        Else
            MsgBox(Me.Page, "Conta nao Permite Cliente.", eTitulo.Info)
        End If
    End Sub

End Class