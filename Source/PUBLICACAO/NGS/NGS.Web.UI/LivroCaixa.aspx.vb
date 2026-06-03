Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LivroCaixa
    Inherits BasePage

    Private dateEmissao As New DateTime
    Private intPagina As Integer = 0
    Dim Cliente As String
    Dim campo() As String
    Private strMoeda As String = "O"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("LivroCaixa", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")
                HID.Value = Guid.NewGuid().ToString()
                ucConsultaPlanoDeContas.SetarHID(HID.Value)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        Try
            If Session("objConsultaPlanoDeContasLC" & HID.Value) IsNot Nothing Then
                Dim objConta = CType(HttpContext.Current.Session("objConsultaPlanoDeContasLC" & HID.Value), [Lib].Negocio.PlanoDeConta)
                txtConta.Text = objConta.Conta & " - " & objConta.Titulo
                Session.Remove("objConsultaPlanoDeContasLC" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(DdlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocio, DdlEmpresa)
    End Sub

    Protected Sub DdlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocio.SelectedValue.ToString(), True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If DdlEmpresa.SelectedIndex = -1 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf txtData.Text.Length = 0 Or IsDate(txtData.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida.")
            txtData.Focus()
            Return False
        ElseIf txtConta.Text = "" Then
            MsgBox(Me.Page, "Informe uma conta.")
            txtConta.Focus()
            Return False
        ElseIf txtNumero.Text = "" Then
            MsgBox(Me.Page, "Informe o número.")
            txtConta.Focus()
            Return False
        End If

        Return True
    End Function

    Protected Sub btnCodigoConta_Click(sender As Object, e As EventArgs) Handles btnCodigoConta.Click
        Try
            HttpContext.Current.Session("ssCampo") = "LivroCaixa"
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me.Page, "objConsultaPlanoDeContasLC" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "Eventos para seleção do número do caixa"

    Protected Sub txtNumero_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNumero.TextChanged
        Try
            If ValidarCampos() Then
                If txtNumero.Text <> "" Then
                    MostrarValores()
                    ValidaValores()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Atualização dos dados"

    Public Function IncluirCaixa() As Boolean
        Dim strSQL As String
        Dim alSQL As New ArrayList

        Cliente = DdlEmpresa.SelectedValue
        campo = Cliente.Split("-")
        Dim Empresa As String = campo(0)
        Dim EndEmpresa As String = campo(1)
        Dim Conta() As String = txtConta.Text.Split("-")

        strSQL = "INSERT INTO Caixa " & vbCrLf & _
                 "(Empresa_Id, EndEmpresa_Id, Conta_Id, Movimento_Id, Caixa_Id, " & vbCrLf & _
                 "Dinheiro, Cheques, Valoes, Selos, Outros) " & vbCrLf & _
                 "VALUES ('" & Empresa & "', " & EndEmpresa & ", " & vbCrLf & _
                 "'" & Conta(0).Trim() & "', " & vbCrLf & _
                 "'" & txtData.Text.ToSqlDate() & "', " & vbCrLf & _
                 txtNumero.Text & ", " & Str(CDec(txtDinheiro.Text)) & ", " & vbCrLf & _
                 Str(CDec(txtCheques.Text)) & ", " & Str(CDec(txtVales.Text)) & ", " & vbCrLf & _
                 Str(CDec(txtSelos.Text)) & ", " & Str(CDec(txtOutros.Text)) & ")" & vbCrLf

        alSQL.Add(strSQL)
        If Banco.GravaBanco(alSQL) = True Then
            'Dim objCaixa(6) As Object
            'objCaixa(0) = txtNumero.Text
            'objCaixa(1) = txtData.Text
            'objCaixa(2) = txtDinheiro.Text
            'objCaixa(3) = txtCheques.Text
            'objCaixa(4) = txtVales.Text
            'objCaixa(5) = txtSelos.Text
            'objCaixa(6) = txtOutros.Text
            'dsCaixa.Tables("Caixa").Rows.Add(objCaixa)
            Return True
        Else
            Return False
        End If
    End Function

#End Region

#Region "Resultados"

    Private Sub MostrarValores()
        If txtNumero.Text.Length > 0 Then
            If Trim(txtSaldoAnterior.Text) = "" Then txtSaldoAnterior.Text = 0
            Dim dsSaldoAnt As DataSet = CompletarCampos(GetSQLSaldoAnterior(), GetCamposSaldoAnterior(), "Razao")
            Dim dsMovimento As DataSet = CompletarCampos(GetSQLMovimentacaoDia(), GetCamposMovimentacaoDia(), "Razao")
            Dim dsCaixa As DataSet = Banco.ConsultaDataSet(GetSQLCaixa(), "Caixa")

            If dsCaixa.Tables(0).Rows.Count = 0 Then

                If String.IsNullOrWhiteSpace(txtSaldoAnterior.Text) Then txtSaldoAnterior.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtSaidasDia.Text) Then txtSaidasDia.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtEntradasDia.Text) Then txtEntradasDia.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtSaldoAtual.Text) Then txtSaldoAtual.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtDinheiro.Text) Then txtDinheiro.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtCheques.Text) Then txtCheques.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtVales.Text) Then txtVales.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtSelos.Text) Then txtSelos.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtOutros.Text) Then txtOutros.Text = "0,00"
                If String.IsNullOrWhiteSpace(txtTotal.Text) Then txtTotal.Text = "0,00"

                'If Not String.IsNullOrWhiteSpace(txtNumero.Text) AndAlso CInt(txtNumero.Text) = 1 Then
                '    BloquearCamposValor(False)
                '    lnkRelatorio.Enabled = True
                'Else
                BloquearCamposValor(True)
                txtDinheiro.Focus()
                'End If
            Else
                txtDinheiro.Text = dsCaixa.Tables(0).Rows(0).Item(2)
                txtCheques.Text = dsCaixa.Tables(0).Rows(0).Item(3)
                txtVales.Text = dsCaixa.Tables(0).Rows(0).Item(4)
                txtSelos.Text = dsCaixa.Tables(0).Rows(0).Item(5)
                txtOutros.Text = dsCaixa.Tables(0).Rows(0).Item(6)
                txtTotal.Text = CDec(txtDinheiro.Text) + CDec(txtCheques.Text) + CDec(txtVales.Text) + CDec(txtSelos.Text) + CDec(txtOutros.Text)
                BloquearCamposValor(False)
                lnkRelatorio.Enabled = True
            End If

            txtSaldoAtual.Text = CDec(txtSaldoAnterior.Text) + CDec(txtEntradasDia.Text) - CDec(txtSaidasDia.Text)
            If CDec(txtSaldoAtual.Text) < 0 Then
                MsgBox(Me.Page, "Não é possível imprimir este caixa, pois o saldo atual é negativo.")
                lnkRelatorio.Enabled = False
            ElseIf CDec(txtSaldoAtual.Text) = 0 Then
                If txtEntradasDia.Text > 0 Or txtSaidasDia.Text > 0 Then
                    lnkRelatorio.Enabled = True
                Else
                    MsgBox(Me.Page, "Não é possível imprimir este caixa, não existe movimento para a data informada.")
                    lnkRelatorio.Enabled = False
                End If
            Else
                txtSaldoAtual.Text = Format(CDec(txtSaldoAtual.Text), "Standard")
            End If
        End If
    End Sub

    Private Function CompletarCampos(ByVal SQL As String, ByVal Campos As ArrayList, ByVal Tabela As String) As DataSet
        Dim intPos As Integer
        Dim dsCompletar As DataSet = Banco.ConsultaDataSet(SQL, Tabela)

        If dsCompletar.Tables(0).Rows.Count = 1 Then
            Select Case dsCompletar.Tables(Tabela).Rows.Count
                Case Is = 1
                    For intPos = 0 To (Campos.Count - 1)
                        Campos.Item(intPos).Text = dsCompletar.Tables(Tabela).Rows(0)(intPos).ToString()
                    Next
            End Select
        End If

        Return dsCompletar
    End Function

    Private Function GetCamposSaldoAnterior() As ArrayList
        Dim alCampos As New ArrayList
        alCampos.Add(txtSaldoAnterior)
        Return alCampos
    End Function

    Private Function GetCamposMovimentacaoDia() As ArrayList
        Dim alCampos As New ArrayList
        alCampos.Add(txtEntradasDia)
        alCampos.Add(txtSaidasDia)
        Return alCampos
    End Function

    Public Sub BloquearCamposValor(ByVal acessar As Boolean)
        Try
            txtDinheiro.Enabled = acessar
            txtCheques.Enabled = acessar
            txtVales.Enabled = acessar
            txtSelos.Enabled = acessar
            txtOutros.Enabled = acessar
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "SQL"

    Private Function GetSQLSaldoAnterior() As String
        Dim strSQL As String
        Cliente = DdlEmpresa.SelectedValue
        campo = Cliente.Split("-")
        Dim Empresa As String = campo(0)
        Dim EndEmpresa As String = campo(1)
        Dim Conta() As String = txtConta.Text.Split("-")

        strSQL = "SELECT ISNULL(SUM(COALESCE(Debito" & IIf(strMoeda = "O", "Oficial", "Moeda") & ", 0) - COALESCE(Credito" & IIf(strMoeda = "O", "Oficial", "Moeda") & ", 0)),0) AS SaldoAnt " & vbCrLf & _
                 "FROM Razao " & vbCrLf & _
                 "WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                 "AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf & _
                 "AND Conta_Id = '" & Conta(0).Trim() & "' " & vbCrLf & _
                 "AND Movimento_Id < '" & txtData.Text.ToSqlDate() & "' " & vbCrLf

        Return strSQL
    End Function

    Private Function GetSQLMovimentacaoDia() As String
        Dim strSQL As String
        Cliente = DdlEmpresa.SelectedValue
        campo = Cliente.Split("-")
        Dim Empresa As String = campo(0)
        Dim EndEmpresa As String = campo(1)
        Dim Conta() As String = txtConta.Text.Split("-")

        strSQL = "SELECT COALESCE(SUM(Debito" & IIf(strMoeda = "O", "Oficial", "Moeda") & "), 0) AS Entradas, " & vbCrLf & _
                 "COALESCE(SUM(Credito" & IIf(strMoeda = "O", "Oficial", "Moeda") & "), 0) AS Saidas " & vbCrLf & _
                 "FROM Razao R " & vbCrLf & _
                 "WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                 "AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf & _
                 "AND Conta_Id = '" & Conta(0).Trim() & "' " & vbCrLf & _
                 "AND Movimento_Id = '" & txtData.Text.ToSqlDate() & "' " & vbCrLf

        Return strSQL
    End Function

    Private Function GetSQLRazao() As String
        Dim strSQL As String
        Cliente = DdlEmpresa.SelectedValue
        campo = Cliente.Split("-")
        Dim Empresa As String = campo(0)
        Dim EndEmpresa As String = campo(1)
        Dim Conta() As String = txtConta.Text.Split("-")

        strSQL = "SELECT Lote_Id, Sequencia_Id, Historico, Debito" & IIf(strMoeda = "O", "Oficial", "Moeda") & " AS DebitoOficial, Credito" & IIf(strMoeda = "O", "Oficial", "Moeda") & " AS CreditoOficial " & vbCrLf & _
                 "FROM Razao R " & vbCrLf & _
                 "WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                 "AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf & _
                 "AND Conta_Id = '" & Conta(0).Trim() & "' " & vbCrLf & _
                 "AND Movimento_Id = '" & txtData.Text.ToSqlDate() & "' " & vbCrLf & _
                 "UNION " & vbCrLf & _
                 "SELECT '' as Lote_Id, '' as Sequencia_Id, '' as Historico, 0  AS DebitoOficial, 0 AS CreditoOficial " & vbCrLf

        Return strSQL
    End Function

    Private Function GetSQLCaixa() As String
        Dim strSQL As String
        Cliente = DdlEmpresa.SelectedValue
        campo = Cliente.Split("-")
        Dim Empresa As String = campo(0)
        Dim EndEmpresa As String = campo(1)
        Dim Conta() As String = txtConta.Text.Split("-")

        strSQL = "SELECT Caixa_Id, Movimento_Id, Dinheiro, Cheques, Valoes, Selos, Outros " & vbCrLf & _
                 "FROM Caixa " & vbCrLf & _
                 "WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                 "AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf & _
                 "AND Conta_Id = '" & Conta(0).Trim() & "' " & vbCrLf & _
                 "AND Movimento_Id = '" & txtData.Text.ToSqlDate() & "' " & vbCrLf & _
                 "AND Caixa_Id = " & txtNumero.Text & " " & vbCrLf

        Return strSQL
    End Function

#End Region

    Private Sub ValoresInformados()
        If ValidaValores() Then
            Try
                txtTotal.Text = Format(CDec(txtDinheiro.Text) + CDec(txtCheques.Text) + CDec(txtVales.Text) + CDec(txtSelos.Text) + CDec(txtOutros.Text), "Standard")
                lnkRelatorio.Enabled = (CDec(txtTotal.Text) = CDec(txtSaldoAtual.Text))
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        End If
    End Sub

    Private Sub Limpar()
        If Not Session("Caixa") Is Nothing Then Session.Remove("Caixa")
        If Not Session("ContaLivroCaixa") Is Nothing Then Session.Remove("ContaLivroCaixa")
        txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")

        txtConta.Text = String.Empty

        txtNumero.Text = String.Empty
        txtSaldoAnterior.Text = String.Empty
        txtSaidasDia.Text = String.Empty
        txtEntradasDia.Text = String.Empty
        txtSaldoAtual.Text = String.Empty

        txtDinheiro.Text = String.Empty
        txtCheques.Text = String.Empty
        txtVales.Text = String.Empty
        txtSelos.Text = String.Empty
        txtOutros.Text = String.Empty
        txtTotal.Text = String.Empty

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaPlanoDeContas.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocio.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Function ValidaValores()
        If txtSaldoAnterior.Text = "" Then
            txtSaldoAnterior.Text = "0,00"
        End If

        If txtSaidasDia.Text = "" Then
            txtSaidasDia.Text = "0,00"
        End If

        If txtEntradasDia.Text = "" Then
            txtEntradasDia.Text = "0,00"
        End If

        If txtSaldoAtual.Text = "" Then
            txtSaldoAtual.Text = "0,00"
        End If

        If txtDinheiro.Text = "" Then
            txtDinheiro.Text = "0,00"
        End If

        If txtCheques.Text = "" Then
            txtCheques.Text = "0,00"
        End If

        If txtVales.Text = "" Then
            txtVales.Text = "0,00"
        End If

        If txtSelos.Text = "" Then
            txtSelos.Text = "0,00"
        End If

        If txtOutros.Text = "" Then
            txtOutros.Text = "0,00"
        End If

        If txtTotal.Text = "" Then
            txtTotal.Text = "0,00"
        End If

        Return True
    End Function

    Public Sub BuscarRegistros(ByVal dsCaixa As DataSet, ByVal dsRazao As DataSet, ByVal dsSaldoAntAux As DataSet)
        Try
            AlimentaCrptRelatorios(dsCaixa, dsRazao, dsSaldoAntAux, "~/Reports/cr_LivroCaixa")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal ds As DataSet, ByVal Ds2 As DataSet, ByVal Ds3 As DataSet, ByVal Caminho As String)

        Dim crptRelatorio As New ReportDocument()

        Try
            Dim dsLivroCaixa As New DataSet
            With dsLivroCaixa
                .Merge(ds)
                .Merge(Ds2)
                .Merge(Ds3)
            End With

            crptRelatorio.FileName = Server.MapPath(Caminho & ".rpt")
            crptRelatorio.Load(crptRelatorio.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
            Dim arquivo As String = NomeArquivo

            crptRelatorio.SetDataSource(dsLivroCaixa)

            Dim crParameterValues As CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As CrystalDecisions.Shared.ParameterDiscreteValue
            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            crParameterFieldDefinitions = crptRelatorio.DataDefinition.ParameterFields()

            If Caminho = "~/Reports/cr_LivroCaixa" Then
                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Nome")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cidade")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Titulo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "MOVIMENTO DO CAIXA No. " & dsLivroCaixa.Tables(0).Rows(0).Item("Caixa_Id") & "/" & Format(dsLivroCaixa.Tables(0).Rows(0).Item("Movimento_Id"), "yyyy") & ""
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Pagina")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Página"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("DataMovimento")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Movimento do Dia: " & Format(dsLivroCaixa.Tables(0).Rows(0).Item("Movimento_Id"), "dd/MM/yyyy")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("LoteSeq")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Lote   Seq"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Descricao")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Descrição"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Entradas")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Entradas"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Saidas")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saídas"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("DetalhesSaldo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Detalhes do Saldo"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Dinheiro")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Dinheiro..."
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cheques")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Cheques...."
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Vales")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Vales......"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Selos")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Selos......"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Outros")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Outros......"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("TotaisDia")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Totais do Dia......"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("SaldoAnterior")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saldo Anterior......"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("SaldoAtual")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saldo Atual......"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("FechamentoDia")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Fechamento do Dia..."
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Caixa")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Caixa"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Visto")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Visto"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            End If

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            If IO.File.Exists(arquivo) Then
                Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Protected Sub txtDinheiro_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDinheiro.TextChanged
        Try
            ValoresInformados()
            txtCheques.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtCheques_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ValoresInformados()
            txtVales.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtVales_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ValoresInformados()
            txtSelos.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtSelos_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ValoresInformados()
            txtOutros.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtOutros_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOutros.TextChanged
        Try
            ValoresInformados()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidarCampos() Then
                Dim blnListar As Boolean = True
                Dim dsRazao As DataSet
                Dim dsSaldoAntAux As DataSet = New DataSet
                Dim dsCaixa As DataSet = Banco.ConsultaDataSet(GetSQLCaixa(), "Caixa")
                Dim dsSaldoAnt As DataSet = CompletarCampos(GetSQLSaldoAnterior(), GetCamposSaldoAnterior(), "Razao")

                If dsCaixa.Tables("Caixa").Rows.Count = 0 Then
                    blnListar = IncluirCaixa()
                    dsCaixa = Banco.ConsultaDataSet(GetSQLCaixa(), "Caixa")
                End If

                If blnListar = True Then
                    dsRazao = Banco.ConsultaDataSet(GetSQLRazao(), "Razao")
                    With dsSaldoAntAux
                        .Tables.Add("SaldoAnterior")
                        With .Tables("SaldoAnterior")
                            .Columns.Add("SaldoAnt")
                            If IsDBNull(dsSaldoAnt.Tables("Razao").Rows(0).Item("SaldoAnt")) = True Then
                                Dim objRegistro() As Object = New Object(0) {0}
                                .Rows.Add(objRegistro)
                            Else : .Rows.Add(dsSaldoAnt.Tables("Razao").Rows(0).ItemArray)
                            End If
                        End With
                    End With

                    txtNumero.Text = ""
                    txtSaldoAnterior.Text = ""
                    txtSaidasDia.Text = ""
                    txtEntradasDia.Text = ""
                    txtSaldoAtual.Text = ""
                    txtDinheiro.Text = ""
                    txtCheques.Text = ""
                    txtVales.Text = ""
                    txtSelos.Text = ""
                    txtOutros.Text = ""
                    txtTotal.Text = ""
                    BuscarRegistros(dsCaixa, dsRazao, dsSaldoAntAux)
                End If
            End If
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LivroCaixa")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class