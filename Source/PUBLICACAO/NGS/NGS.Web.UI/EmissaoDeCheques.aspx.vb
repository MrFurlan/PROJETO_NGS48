Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Printing
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.ServiceModel

Public Class EmissaoDeCheques
    Inherits BasePage

    Dim linha As String
    Dim cl As Integer = 0
    Dim pagina As Integer = 0
    Dim Sql As String
    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim Codigo As String
    Dim Descricao As String
    Dim Cliente As String
    Dim Endereco As String
    Dim index As Integer

#Region "Métodos"

    Private Sub BuscarEmpresas()
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
    End Sub

    Private Sub BuscarUnidadesDeNegocio()
        ddl.Carregar(DdlUnidadeDeNegocioEmpresaCliente, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidadeDeNegocio()
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocioEmpresaCliente, DdlEmpresaCliente)
    End Sub

    Private Sub CargaBancos()
        ddl.Carregar(DdlBancoPagador, CarregarDDL.Tabela.Bancos, "", True)
    End Sub

    Private Sub CarregarImpressoras()
        Dim fwPrintSetting As New System.Drawing.Printing.PrinterSettings
        Dim nCnt As Integer

        With ddlImpressora.Items
            For nCnt = 0 To (PrinterSettings.InstalledPrinters.Count - 1)
                .Add(PrinterSettings.InstalledPrinters.Item(nCnt))
            Next
        End With

        Funcoes.InserirLinhaEmBranco(ddlImpressora)

        Sql = "Select Descricao, IpPrinter FROM RotinasDeImpressao WHERE Rotina_Id = 'CHEQUE'"

        Dim dsImpressora = Banco.ConsultaDataSet(Sql, "Impressora")

        If dsImpressora Is Nothing Then

        ElseIf dsImpressora.Tables(0).Rows.Count = 0 Then

        Else
            HttpContext.Current.Session("printerFinanceiro") = dsImpressora.Tables(0).Rows(0).Item("Descricao") & "|" & dsImpressora.Tables(0).Rows(0).Item("IpPrinter")
            With ddlImpressora
                .SelectedIndex = .Items.IndexOf(.Items.FindByValue(dsImpressora.Tables(0).Rows(0).Item("Descricao")))
            End With
        End If

    End Sub

    Private Sub BancoPagador()
        Dim Conta As String

        DdlContaPagadora.Items.Clear()

        If DdlEmpresaCliente.SelectedIndex > 0 Then
            Sql = " SELECT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes " & vbCrLf & _
                  "   FROM BancosXContas " & vbCrLf & _
                  "  WHERE Empresa_Id = '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id  = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & " and " & vbCrLf
        Else
            Sql = "SELECT DISTINCT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes " & vbCrLf & _
                  " FROM BancosXContas Where " & vbCrLf
        End If
        Sql &= " Banco_Id  = " & DdlBancoPagador.SelectedValue      'Endereco da Empresa

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
            Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("ContaContabil")
            Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "  C/C " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes")
            DdlContaPagadora.Items.Add(New ListItem(Descricao, Conta))
        Next

        DdlContaPagadora.Items.Insert(0, "")
        DdlContaPagadora.SelectedIndex = 0
    End Sub

    Private Function ValidarCampos(ByVal validaEmissao As Boolean) As Boolean
        If DdlEmpresaCliente.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Informe a Empresa!")
            Return False
        ElseIf DdlBancoPagador.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Informe o banco!")
            Return False
        ElseIf DdlContaPagadora.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Informe a agência!")
            Return False
            'ElseIf ddlImpressora.SelectedIndex = 0 Then
            '    MsgBox(Me.Page, "Impressora não foi selecionada!")
            '    Return False
        ElseIf validaEmissao AndAlso txtChequeinicial.Text = 0 Then
            MsgBox(Me.Page, "Não existem Folhas de cheques cadastrado para esta conta!")
            Return False
        ElseIf validaEmissao Then
            Return validaFolhasDeCheque()
        End If

        Return True
    End Function

    Private Sub Limpar()
        lnkEmitir.Parent.Visible = False
        DdlBancoPagador.SelectedIndex = 0
        DdlContaPagadora.Items.Clear()
        txtMovimento.Text = Format(CDate(Today), "dd/MM/yyyy")
        chkReimprime.Checked = False
        txtChequeinicial.Text = String.Empty
        txtChequeinicial.Enabled = False
        ddlImpressora.SelectedIndex = 0
        GridConsultaTitulos1.DataSource = New List(Of Object)
        GridConsultaTitulos1.DataBind()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
            DdlEmpresaCliente.Enabled = False
        End If
    End Sub

    Private Sub getNumCheque()
        If Not String.IsNullOrWhiteSpace(DdlEmpresaCliente.SelectedValue) AndAlso Not String.IsNullOrWhiteSpace(DdlBancoPagador.SelectedValue) AndAlso Not String.IsNullOrWhiteSpace(DdlContaPagadora.SelectedValue) Then
            Dim objBancoXConta As New [Lib].Negocio.BancosXContas(DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), DdlBancoPagador.SelectedValue.Split("-")(0), _
                                                                         DdlContaPagadora.SelectedValue.Split("-")(0), DdlContaPagadora.SelectedValue.Split("-")(1), DdlContaPagadora.SelectedValue.Split("-")(2), DdlContaPagadora.SelectedValue.Split("-")(3))

            If objBancoXConta.NumChequeAtual = objBancoXConta.NumChequeFinal Then
                MsgBox(Me.Page, "Não existem folhas de cheques cadastrados para emissões.")
                txtChequeinicial.Text = 0
            ElseIf objBancoXConta.NumChequeAtual = 0 Then
                txtChequeinicial.Text = objBancoXConta.NumChequeInicial
            Else
                txtChequeinicial.Text = objBancoXConta.NumChequeAtual + 1
            End If

            'Se o cheque estiver Inutilizado pegar próximo
            Dim ListBXCXC As New [Lib].Negocio.ListBancosXContasXCancelamentoCheque(objBancoXConta, CInt(txtChequeinicial.Text) - 1)

            For Each objBXCXC As [Lib].Negocio.BancosXContasXCancelamentoCheque In ListBXCXC
                If objBXCXC.NumCheque = txtChequeinicial.Text Then
                    txtChequeinicial.Text = objBXCXC.NumCheque + 1
                End If
            Next
        End If
    End Sub

    Protected Sub CargaGridconsultaTitulos()
        Try
            Dim ds As New DataSet
            Dim Cliente As String
            Dim Campo() As String
            Dim Unidade As String = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue

            Sql = " SELECT ContasAPagar.Registro_id AS Registro, ContasAPagar.Sequencia_id AS Sequencia, ContasAPagar.Lote, " & vbCrLf & _
                  "        Clientes.Nome AS destinatario, ContasAPagar.Historico, ContasAPagar.ValorLiquido AS Valor, Empresas.Cidade, " & vbCrLf & _
                  "        Empresas.Estado , ContasApagar.Destinacao AS Destinacao, ContasAPagar.ContaContabilPagadora as ContaContabilPagadora " & vbCrLf & _
                  "   FROM ContasAPagar " & vbCrLf & _
                  "  INNER JOIN Clientes " & vbCrLf & _
                  "     ON ContasAPagar.Cliente = Clientes.Cliente_Id " & vbCrLf & _
                  "    AND ContasAPagar.Endcliente = Clientes.Endereco_Id " & vbCrLf & _
                  "  INNER JOIN Clientes AS Empresas" & vbCrLf & _
                  "     ON ContasAPagar.EmpresaPagadora    = Empresas.Cliente_Id " & vbCrLf & _
                  "    AND ContasAPagar.EndEmpresaPagadora = Empresas.Endereco_Id    " & vbCrLf & _
                  "  WHERE ContasAPagar.Situacao  = 1" & vbCrLf & _
                  "    and ContasAPagar.Provisao  = 1" & vbCrLf & _
                  "    and ContasAPagar.TipoPagto = 2" & vbCrLf & _
                  "    and ContasAPagar.Grupado IN ('N','M') " & vbCrLf

            If chkReimprime.Checked = False Then
                Sql &= " and ContasAPagar.Cheque = 'N'"
            Else
                Sql &= " and ContasAPagar.Cheque = 'S'"
            End If

            If DdlEmpresaCliente.SelectedIndex > 0 Then
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")
                If Campo(0) <> "" Then
                    Sql &= " and ContasAPagar.EmpresaPagadora = '" & Campo(0) & "'" & vbCrLf & _
                           " and ContasAPagar.EndEmpresaPagadora = " & Campo(1) & vbCrLf
                End If
            End If

            Dim Banco1 As String = DdlBancoPagador.SelectedValue

            If Banco1 <> "" Then
                Sql &= " and ContasAPagar.BancoPagador = '" & Banco1 & "'"   'Banco
            End If

            Dim Agencia As String = DdlContaPagadora.SelectedValue
            Campo = Agencia.Split("-")
            If Campo(0) <> "" Then
                Sql &= " and ContasAPagar.AgenciaPagadora = '" & Campo(0) & "'" & vbCrLf & _
                       " and ContasAPagar.DigitoAgenciaPagadora = '" & Campo(1) & "'" & vbCrLf & _
                       " and ContasAPagar.ContaPagadora = '" & Campo(2) & "'" & vbCrLf & _
                       " and ContasAPagar.DigitoContaPagadora = '" & Campo(3) & "'" & vbCrLf
            End If

            Dim Data1 As String = txtMovimento.Text.ToSqlDate()

            If Data1 <> "" Then
                Sql &= " and ContasAPagar.baixa = '" & Data1 & "'"   'Data de baixa
            End If

            Sql &= " ORDER BY Clientes.Nome"

            ds = Banco.ConsultaDataSet(Sql, "Contas1")

            If ds.Tables(0).Rows.Count > 0 Then
                GridConsultaTitulos1.DataSource = ds
                GridConsultaTitulos1.DataBind()

                For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
                    If CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Text.Length = 0 Then
                        CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Enabled = True
                    End If
                Next intPos

                lnkEmitir.Parent.Visible = True
            Else
                MsgBox(Me.Page, "Nenhum resultado encontrado.")
                GridConsultaTitulos1.DataSource = New List(Of Object)
                GridConsultaTitulos1.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaGridConsultaTituloNovo()
        Dim ds As New DataSet

        Sql = "   Select distinct r.Lote_Id as Lote, t.Titulo_Id as Registro, isnull(cli.Nome,'') as Destinatario, t.Historico, emp.Cidade,                                                     " & vbCrLf & _
              "   		emp.Estado, td.Destinacao, ContaContabilRecPag as ContaContabilPagadora, '' as Sequencia,                                                                             " & vbCrLf & _
              "   	(Select SUM(case                                                                                                                                                          " & vbCrLf & _
              "	  when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end    " & vbCrLf & _
              "	  when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end    " & vbCrLf & _
              "	  else 0                                                                                                                                                                      " & vbCrLf & _
              "        End                                                                                                                                                                      " & vbCrLf & _
              "	)                                                                                                                                                                             " & vbCrLf & _
              "	from Titulos Tp                                                                                                                                                               " & vbCrLf & _
              "	inner Join TitulosxContaContabil Tc                                                                                                                                           " & vbCrLf & _
              "	on Tc.Titulo_Id = Tp.Titulo_Id                                                                                                                                                " & vbCrLf & _
              "	INNER JOIN Moedas M                                                                                                                                                           " & vbCrLf & _
              "	on M.Moeda_id = Tp.Moeda                                                                                                                                                      " & vbCrLf & _
              "	where tp.Titulo_Id = t.Titulo_Id                                                                                                                                              " & vbCrLf & _
              "	Group by Tp.Titulo_Id) as Valor                                                                                                                                       " & vbCrLf & _
              "   	From	Titulos t                                                                                                                                                         " & vbCrLf & _
              "   		left Join TituloXDestinacao td                                                                                                                                       " & vbCrLf & _
              "   			On td.Titulo_Id = t.Titulo_Id                                                                                                                                     " & vbCrLf & _
              "   		Left Join Clientes cli                                                                                                                                                " & vbCrLf & _
              "   			On	cli.Cliente_Id	=  t.CliFor                                                                                                                                   " & vbCrLf & _
              "   			And cli.Endereco_Id = t.EnderecoCliFor                                                                                                                            " & vbCrLf & _
              "   		Left Join Clientes emp                                                                                                                                                " & vbCrLf & _
              "   			On	emp.Cliente_Id	=  t.EmpresaRecPag                                                                                                                            " & vbCrLf & _
              "   			And emp.Endereco_Id = t.EndEmpresaRecPag			                                                                                                              " & vbCrLf & _
              "   LEFT JOIN BancosXContas BC                                                                                                                                                    " & vbCrLf & _
              "   			ON BC.ContaContabil = T.ContaContabilRecPag	                                                                                                                      " & vbCrLf & _
              "   		 LEFT JOIN Bancos AS B		                                                                                                                                          " & vbCrLf & _
              "   			ON B.Banco_Id = BC.Banco_Id                                                                                                                                       " & vbCrLf & _
              "   Inner Join Razao r                                                                                                                                                            " & vbCrLf & _
              "          		on r.Titulo = t.Titulo_Id                                                                                                                                         " & vbCrLf & _
              "   	Where		t.Situacao	= 1                                                                                                                                               " & vbCrLf & _
              "   			And t.Provisao	= 1                                                                                                                                               " & vbCrLf & _
              "   			And t.RecPag	= 'P'                                                                                                                                             " & vbCrLf & _
              "   			And (t.RegistroMestre is null or t.RegistroMestre = t.Titulo_Id)                                                                                                  " & vbCrLf

        If chkReimprime.Checked Then
            Sql &= " And t.Cheque = 1 " & vbCrLf
        Else
            Sql &= " and t.Cheque = 0 " & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlUnidadeDeNegocioEmpresaCliente.SelectedValue) Then
            Sql &= " and t.UnidadeDeNegocioRecPag = '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "' " & vbCrLf ' Unidade 
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresaCliente.SelectedValue) Then
            Sql &= " and t.EmpresaRecPag = '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                " and t.EndEmpresaRecPag = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & vbCrLf     'Endereco da Empresa
        End If
        If Not String.IsNullOrWhiteSpace(DdlBancoPagador.SelectedValue) Then
            Sql &= " and BC.Banco_Id = '" & DdlBancoPagador.SelectedValue & "'" & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlContaPagadora.SelectedValue) Then
            Sql &= " and BC.Agencia_Id = '" & DdlContaPagadora.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                " and BC.DigitoAgencia_Id = '" & DdlContaPagadora.SelectedValue.Split("-")(1) & "'" & vbCrLf & _
                " and BC.Conta_Id = '" & DdlContaPagadora.SelectedValue.Split("-")(2) & "'" & vbCrLf & _
                " and BC.DigitoConta_Id = '" & DdlContaPagadora.SelectedValue.Split("-")(3) & "'" & vbCrLf
        End If
        If IsDate(txtMovimento.Text) Then
            Sql &= " and t.DataBaixa = '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf
        End If

        Sql &= " ORDER BY isnull(cli.Nome,'')" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Contas1")
        GridConsultaTitulos1.DataSource = ds
        GridConsultaTitulos1.DataBind()

        For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
            If CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Text.Length = 0 Then
                CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Enabled = True
            End If
        Next intPos

        If GridConsultaTitulos1.Rows.Count > 0 Then lnkEmitir.Parent.Visible = True
    End Sub

#End Region

#Region "Eventos"

    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EmissaoDeCheques", "ACESSAR") Then
                CarregarImpressoras()
                CargaBancos()
                Limpar()
                BuscarUnidadesDeNegocio()
                VerificaUnidadeDeNegocio()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscarEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlBancoPagador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlBancoPagador.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(DdlBancoPagador.SelectedValue) Then
                If DdlBancoPagador.SelectedValue = 1 OrElse _
                    DdlBancoPagador.SelectedValue = 756 OrElse _
                    DdlBancoPagador.SelectedValue = 1001 Then
                    BancoPagador()
                Else
                    MsgBox(Me.Page, "Só estão liberados para emissão de cheques os Bancos 1 - BANCO DO BRASIL, 756 - BANCOOB/SICOOB e 1001 - PRIMACREDI.")
                    DdlContaPagadora.Items.Clear()
                End If
            Else
                DdlContaPagadora.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ValidarCampos(False) Then
                If FinanceiroNovo Then
                    CargaGridConsultaTituloNovo()
                Else
                    CargaGridconsultaTitulos()
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

    Private Function validaFolhasDeCheque() As Boolean
        Try
            Dim selecionados As Integer

            For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
                If CType(GridConsultaTitulos1.Rows(intPos).FindControl("ChkImprimir"), CheckBox).Checked = True Then selecionados += 1
            Next intPos

            Dim objBXC As New [Lib].Negocio.BancosXContas(DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), _
                                                               DdlBancoPagador.SelectedValue, DdlContaPagadora.SelectedValue.Split("-")(0), _
                                                               DdlContaPagadora.SelectedValue.Split("-")(1), DdlContaPagadora.SelectedValue.Split("-")(2), _
                                                               DdlContaPagadora.SelectedValue.Split("-")(3))

            If objBXC.NumChequeAtual = objBXC.NumChequeFinal Then
                MsgBox(Me.Page, "Não existem folhas de cheques cadastrados para emissão.")
                Return False
            ElseIf selecionados = 0 Then
                MsgBox(Me.Page, "Selecione os Titulos para emissão")
                Return False
            ElseIf selecionados > (objBXC.NumChequeFinal - (CInt(txtChequeinicial.Text) - 1)) Then
                MsgBox(Me.Page, "A quantidade de Titulos selecionados é maior que a quantidade de folhas de cheque cadastradas.")
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Protected Sub lnkEmitir_Click(sender As Object, e As EventArgs) Handles lnkEmitir.Click
        Try

            If ddlImpressora.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Impressora não foi selecionada!")
                Exit Sub
            ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "04854422" OrElse Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "03189063" Then
                EmissaoChequeWS()
            Else
                EmissaoCheque()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("EmissaoDeCheques", "RELATORIO") Then
                ImprimirCheque()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ImprimirCheque(Optional ByVal Relatorio As Boolean = True)
        Try
            'If Funcoes.VerificaPermissao("EmissaoDeCheques", "RELATORIO") Then
            ''*** rotina nova do relatorio 
            Dim xextenso As String = String.Empty
            Dim yextenso As String = String.Empty
            Dim zextenso As String = String.Empty
            Dim dsEmitir As New DataSet
            Dim dtEmitir As DataTable
            Dim row As DataRow
            Dim i As Integer
            Dim j As Integer
            Dim k As Integer
            Dim MESX As String = String.Empty
            Dim NumeroDoCheque As Integer
            Dim NumeroDoChequeX As String = String.Empty
            Dim Historico As String = String.Empty
            Dim RegistroI As String = String.Empty
            Dim RegistroS As String = String.Empty

            dtEmitir = New DataTable("Cheques")
            dtEmitir.Columns.Add("Valor", Type.GetType("System.Decimal"))
            dtEmitir.Columns.Add("Portador", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Extenso", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Cidade", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Data", Type.GetType("System.DateTime"))
            dtEmitir.Columns.Add("Destinacao", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Dia", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Mes", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Ano", Type.GetType("System.String"))
            dtEmitir.Columns.Add("NumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
            dtEmitir.Columns.Add("Historico", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Portador1", Type.GetType("System.String"))
            dtEmitir.Columns.Add("Registro", Type.GetType("System.String"))
            dsEmitir.Tables.Add(dtEmitir)

            If txtChequeinicial.Text = "" Then
                NumeroDoCheque = 0
            Else
                NumeroDoCheque = CInt(txtChequeinicial.Text)
            End If

            NumeroDoCheque = NumeroDoCheque - 1
            While i < GridConsultaTitulos1.Rows.Count
                If CType(GridConsultaTitulos1.Rows(i).FindControl("ChkImprimir"), CheckBox).Checked = True Then
                    row = dtEmitir.NewRow()
                    row("Valor") = GridConsultaTitulos1.Rows(i).Cells(5).Text()
                    zextenso = UCase(GridConsultaTitulos1.Rows(i).Cells(3).Text())
                    yextenso = "("
                    yextenso &= UCase(GridConsultaTitulos1.Rows(i).Cells(3).Text())
                    xextenso = yextenso
                    For j = 1 To (79 - Len(xextenso))
                        xextenso &= "*"
                    Next
                    xextenso &= ")"
                    row("Portador") = xextenso
                    yextenso = "("
                    yextenso &= UCase(Funcoes.Extenso(GridConsultaTitulos1.Rows(i).Cells(5).Text(), "Real", "Reais"))
                    yextenso &= " *"
                    xextenso = yextenso
                    For j = 1 To (120 - Len(xextenso))
                        xextenso &= " *"
                    Next
                    xextenso &= ")"
                    row("Extenso") = xextenso
                    row("Cidade") = GridConsultaTitulos1.Rows(i).Cells(7).Text()
                    row("Destinacao") = GridConsultaTitulos1.Rows(i).Cells(8).Text()
                    row("Data") = txtMovimento.Text
                    row("Dia") = Day(CDate(txtMovimento.Text))

                    row("Mes") = MonthName(Month(CDate(txtMovimento.Text))).ToUpper

                    row("Ano") = CStr(Year(CDate(txtMovimento.Text)))
                    NumeroDoCheque = NumeroDoCheque + 1
                    NumeroDoChequeX = CStr(NumeroDoCheque)
                    Dim x As String
                    x = (CStr(Int(Val(NumeroDoCheque))))
                    Dim NumeroDochequey As String
                    NumeroDochequey = " "
                    NumeroDochequey = Trim(NumeroDochequey)
                    For k = 1 To (6 - Len(NumeroDoChequeX))
                        NumeroDochequey &= "0"
                    Next
                    NumeroDochequey &= NumeroDoChequeX
                    row("NumeroDoCheque") = NumeroDochequey
                    Historico = "CH.NR. " & NumeroDochequey & "  " & UCase(GridConsultaTitulos1.Rows(i).Cells(3).Text())
                    row("Historico") = Historico
                    row("Portador1") = zextenso
                    '' calculo de registro . 
                    RegistroI = GridConsultaTitulos1.Rows(i).Cells(0).Text()
                    RegistroS = " "
                    RegistroS = Trim(RegistroS)
                    If Len(RegistroS) < 6 Then
                        For k = 1 To (6 - Len(Trim(RegistroI)))
                            RegistroS &= "0"
                        Next
                    End If
                    row("Registro") = Trim(RegistroS) & (Trim(RegistroI))
                    dtEmitir.Rows.Add(row)
                End If
                i = i + 1
            End While

            Dim param As New Dictionary(Of String, Object)
            param.Add("Titulo", "Relação de cheques emitidos em :")
            param.Add("Empresa", DdlUnidadeDeNegocioEmpresaCliente.SelectedValue)
            param.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
            param.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))

            If Relatorio Then
                Funcoes.BindReport(Me.Page, dsEmitir, "Cr_RelatorioEmissaoDeCheque", eExportType.PDF, param)
            Else
                Funcoes.BindReport(Me.Page, dsEmitir, "Cr_EmissaoDeChequeGrafico", eExportType.PDF, param)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub DdlEmpresaCliente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlEmpresaCliente.SelectedIndexChanged
        Try
            DdlBancoPagador.SelectedIndex = 0
            DdlContaPagadora.Items.Clear()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlContaPagadora_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlContaPagadora.SelectedIndexChanged
        Try
            If ControlaCheques() Then
                getNumCheque()
                txtChequeinicial.Enabled = False
            Else
                txtChequeinicial.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ControlaCheques() As Boolean
        Dim sql As String = "select isnull(ControlaEmissaoCheque, 0) as ControlaEmissaoCheque from ClientesXEmpresas Where Empresa_Id = '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ControlaEmissao")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            If ds.Tables(0).Rows(0)(0) = 1 Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Function EmissaoChequeWS() As Boolean
        Dim testeUrl As String = ""
        Dim Impressora As String() = HttpContext.Current.Session("printerFinanceiro").ToString().Split("|")
        Dim binding As New BasicHttpBinding()
        Dim ipAddress As String = String.Format("http://{0}/wsCheque/Service1.svc", Impressora(1))
        ' testeUrl = ipAddress
        Dim endpointAddress As New EndpointAddress(ipAddress)
        Dim cheques As New List(Of wsCheque.Cheque)
        Dim ws As New wsCheque.Service1Client(binding, endpointAddress)

        Try
            If ValidarCampos(ControlaCheques()) Then
                Dim Esquerda As Integer = 24
                Dim intNumeroCheque As Integer = 1
                Dim selecionados As Integer = 0
                Dim intPosCheque As Integer = 0
                Dim strHistorico As String = ""
                Dim arrSQL As New ArrayList()
                Dim strSQL As String

                For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
                    If CType(GridConsultaTitulos1.Rows(intPos).FindControl("ChkImprimir"), CheckBox).Checked = True Then selecionados += 1
                Next intPos

                Dim Campo() As String
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")

                Dim strBanco As String = DdlBancoPagador.SelectedValue

                If txtChequeinicial.Text.Length > 0 Then intNumeroCheque = Convert.ToInt32(txtChequeinicial.Text)

                Dim dblValor(selecionados - 1) As Double
                Dim strValorExtenso(selecionados - 1) As String
                Dim strNome(selecionados - 1) As String
                Dim strCidade(selecionados - 1) As String
                Dim dateData(selecionados - 1) As DateTime
                Dim intNumero(selecionados - 1) As Integer
                Dim ContaContabilPagadora As String

                For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
                    If CType(GridConsultaTitulos1.Rows(intPos).FindControl("ChkImprimir"), CheckBox).Checked = True Then

                        With GridConsultaTitulos1.Rows(intPos)
                            dblValor(intPosCheque) = Convert.ToDouble(.Cells(5).Text)

                            If CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Text.Length = 0 Then
                                strNome(intPosCheque) = Textos.RetirarAcentos(.Cells(3).Text)
                            Else
                                strNome(intPosCheque) = Textos.RetirarAcentos(CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Text)
                            End If

                            strCidade(intPosCheque) = Textos.RetirarAcentos(.Cells(7).Text)
                            Dim strData As String() = txtMovimento.Text.Split("/")
                            dateData(intPosCheque) = New DateTime(Convert.ToInt32(strData(2)), Convert.ToInt32(strData(1)), Convert.ToInt32(strData(0)))
                            intNumero(intPosCheque) = intNumeroCheque

                        End With

                        strHistorico = "CH.NR. " & intNumeroCheque.ToString("000000") & "  " & strNome(intPosCheque)
                        ContaContabilPagadora = GridConsultaTitulos1.Rows(intPos).Cells(9).Text

                        If DdlEmpresaCliente.SelectedIndex > 0 Then
                            Dim strEmpresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
                            strSQL = "UPDATE Razao " & _
                                     "SET Historico = '" & strHistorico & "' " & _
                                     "WHERE Empresa_id = '" & strEmpresa(0) & "' " & _
                                     "AND EndEmpresa_id = '" & strEmpresa(1) & "' " & _
                                     "AND Conta_id = '" & ContaContabilPagadora & "' " & _
                                     "AND Titulo = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text
                        Else
                            strSQL = "UPDATE Razao " & _
                                     "SET Historico = '" & strHistorico & "' where " & _
                                     " Conta_id = '" & ContaContabilPagadora & "' " & _
                                     "AND Titulo = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text
                        End If

                        arrSQL.Add(strSQL)

                        '' rotina de update de contas a pagar inicio
                        If FinanceiroNovo Then
                            strSQL = "UPDATE TITULOS" & vbCrLf & _
                                     "       SET CHEQUE     = 1, " & vbCrLf & _
                                     "           NumeroDoCheque = " & intNumeroCheque & vbCrLf & _
                                     "   WHERE RECPAG = 'P'" & vbCrLf & _
                                     "       AND TITULO_ID = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text & vbCrLf & _
                                     "" & vbCrLf
                        Else
                            strSQL = " UPDATE ContasAPagar" & vbCrLf & _
                                     "           SET Cheque = 'S'," & vbCrLf & _
                                     "               NumeroDoCheque = " & intNumeroCheque & vbCrLf & _
                                     "       WHERE Registro_ID = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text & vbCrLf
                        End If

                        arrSQL.Add(strSQL)

                        'ATUALIZA NUMERADOR
                        strSQL = " UPDATE BANCOSXCONTAS" & vbCrLf & _
                                 "    SET NumCheque = " & intNumeroCheque & vbCrLf & _
                                 "  WHERE Empresa_Id       = '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                 "    AND EndEmpresa_Id    = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & vbCrLf & _
                                 "    AND Banco_Id         = " & DdlBancoPagador.SelectedValue & vbCrLf & _
                                 "    AND Agencia_Id       = '" & DdlContaPagadora.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                 "    AND DigitoAgencia_Id = '" & DdlContaPagadora.SelectedValue.Split("-")(1) & "'" & vbCrLf & _
                                 "    AND Conta_Id         = '" & DdlContaPagadora.SelectedValue.Split("-")(2) & "'" & vbCrLf & _
                                 "    AND DigitoConta_Id   = '" & DdlContaPagadora.SelectedValue.Split("-")(3) & "'" & vbCrLf

                        arrSQL.Add(strSQL)

                        Dim c As New wsCheque.Cheque()
                        c.CodBanco = DdlBancoPagador.SelectedValue
                        c.Valor = dblValor(intPosCheque)
                        c.ValorExtenso = Left(HttpContext.Current.Session("ssEmpresa").ToString, 8)
                        c.Destinatario = strNome(intPosCheque)
                        c.Cidade = strCidade(intPosCheque)
                        c.Data = dateData(intPosCheque)
                        c.NumCheque = intNumero(intPosCheque)
                        c.Impressora = Impressora(0)
                        cheques.Add(c)
                        '' rotina de update de contas a pagar fim 
                        intPosCheque += 1
                        intNumeroCheque += 1
                    End If
                Next

                'Arquivo.Close()


                If Banco.GravaBanco(arrSQL) Then
                    ws.EmissaoDeChequeList(cheques)
                    ws.Close()
                    MsgBox(Me.Page, "Os cheques foram emitidos com sucesso! " & testeUrl, eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage")) & ". " & testeUrl)
                End If

                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub EmissaoCheque()
        Dim testeUrl As String = ""
        Try
            If ValidarCampos(ControlaCheques()) Then
                Dim Esquerda As Integer = 24
                Dim intNumeroCheque As Integer = 1
                Dim selecionados As Integer = 0
                Dim intPosCheque As Integer = 0
                Dim strHistorico As String = ""
                Dim arrSQL As New ArrayList()
                Dim strSQL As String

                For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
                    If CType(GridConsultaTitulos1.Rows(intPos).FindControl("ChkImprimir"), CheckBox).Checked = True Then selecionados += 1
                Next intPos

                Dim Campo() As String
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")

                Dim extenso As String
                Dim j As Integer

                Dim Arquivo As System.IO.StreamWriter
                Arquivo = System.IO.File.CreateText(Server.MapPath("~/Files/Cheque.txt"))

                Dim strBanco As String = DdlBancoPagador.SelectedValue

                If txtChequeinicial.Text.Length > 0 Then intNumeroCheque = Convert.ToInt32(txtChequeinicial.Text)

                Dim dblValor(selecionados - 1) As Double
                Dim strValorExtenso(selecionados - 1) As String
                Dim strNome(selecionados - 1) As String
                Dim strCidade(selecionados - 1) As String
                Dim dateData(selecionados - 1) As DateTime
                Dim intNumero(selecionados - 1) As Integer
                Dim ContaContabilPagadora As String

                For intPos As Integer = 0 To GridConsultaTitulos1.Rows.Count - 1
                    If CType(GridConsultaTitulos1.Rows(intPos).FindControl("ChkImprimir"), CheckBox).Checked = True Then
                        With GridConsultaTitulos1.Rows(intPos)
                            dblValor(intPosCheque) = Convert.ToDouble(.Cells(5).Text)
                            strValorExtenso(intPosCheque) = Textos.RetirarAcentos(Funcoes.Extenso(.Cells(5).Text, "Real", "Reais"))

                            extenso = UCase(Funcoes.Extenso(dblValor(intPosCheque).ToString("N2"), "Real", "Reais"))
                            For j = 1 To (170 - Len(extenso))
                                extenso &= "*"
                            Next

                            If CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Text.Length = 0 Then
                                strNome(intPosCheque) = Textos.RetirarAcentos(.Cells(3).Text)
                            Else
                                strNome(intPosCheque) = Textos.RetirarAcentos(CType(GridConsultaTitulos1.Rows(intPos).FindControl("txtDestinacao"), TextBox).Text)
                            End If

                            strCidade(intPosCheque) = Textos.RetirarAcentos(.Cells(7).Text)
                            Dim strData As String() = txtMovimento.Text.Split("/")
                            dateData(intPosCheque) = New DateTime(Convert.ToInt32(strData(2)), Convert.ToInt32(strData(1)), Convert.ToInt32(strData(0)))
                            intNumero(intPosCheque) = intNumeroCheque

                            If DdlBancoPagador.SelectedValue = 756 Then
                                Dim strTexto As String
                                Dim objESCP As New Impressora_ESCP

                                strTexto = objESCP.TamanhoPagina(17) & Impressora_ESCP.PRINT_STYLE_BOLD_OFF
                                strTexto &= Impressora_ESCP.PRINT_CONDENSED_ON
                                strTexto &= Impressora_ESCP.PRINT_CARACTER_10DPI
                                strTexto &= Impressora_ESCP.PRINT_6_LPI
                                strTexto &= Space(Esquerda + 96) & "(**" & dblValor(intPosCheque).ToString("N2") & "**)" & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 16) & Left(extenso, 85) & ControlChars.CrLf
                                strTexto &= Space(Esquerda + 2) & Mid(extenso, 85, 85) & ControlChars.CrLf
                                strTexto &= Space(Esquerda + 2) & strNome(intPosCheque) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 77 - strCidade(intPosCheque).Length) & strCidade(intPosCheque) & Space(2) & strData(0) & Space(7) & Funcoes.EliminarCaracteresEspeciais(MonthName(CDate(dateData(intPosCheque)).Month).ToString.ToUpper()) & Space(21 - Funcoes.EliminarCaracteresEspeciais(MonthName(CDate(dateData(intPosCheque)).Month).ToString.Length)) & strData(2) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 35) & intNumero(intPosCheque) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf

                                Arquivo.WriteLine(strTexto)

                            ElseIf DdlBancoPagador.SelectedValue = 1 Then
                                Dim strTexto As String
                                Dim objESCP As New Impressora_ESCP

                                strTexto = objESCP.TamanhoPagina(23) & Impressora_ESCP.PRINT_STYLE_BOLD_OFF
                                strTexto &= Impressora_ESCP.PRINT_CONDENSED_ON
                                strTexto &= Impressora_ESCP.PRINT_CARACTER_10DPI
                                strTexto &= Impressora_ESCP.PRINT_8_LPI
                                Esquerda = 0
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 87) & "(**" & dblValor(intPosCheque).ToString("N2") & "**)" & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 15) & Left(extenso, 95) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda) & Mid(extenso, 95, 110) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda) & strNome(intPosCheque) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 68 - strCidade(intPosCheque).Length) & strCidade(intPosCheque) & Space(2) & strData(0) & Space(9) & Funcoes.EliminarCaracteresEspeciais(MonthName(CDate(dateData(intPosCheque)).Month).ToString.ToUpper()) & Space(25 - Funcoes.EliminarCaracteresEspeciais(MonthName(CDate(dateData(intPosCheque)).Month).ToString.Length)) & strData(2) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= Space(Esquerda + 35) & intNumero(intPosCheque) & ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf
                                strTexto &= ControlChars.CrLf

                                Arquivo.WriteLine(strTexto)
                            End If
                        End With

                        strHistorico = "CH.NR. " & intNumeroCheque.ToString("000000") & "  " & strNome(intPosCheque)
                        ContaContabilPagadora = GridConsultaTitulos1.Rows(intPos).Cells(9).Text

                        If DdlEmpresaCliente.SelectedIndex > 0 Then
                            Dim strEmpresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
                            strSQL = "UPDATE Razao " & _
                                     "   SET Historico = '" & strHistorico & "' " & _
                                     " WHERE Empresa_id = '" & strEmpresa(0) & "' " & _
                                     "   AND EndEmpresa_id = '" & strEmpresa(1) & "' " & _
                                     "   AND Conta_id = '" & ContaContabilPagadora & "' " & _
                                     "   AND Titulo = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text
                        Else
                            strSQL = "UPDATE Razao " & _
                                     "   SET Historico = '" & strHistorico & "'" & _
                                     " WHERE Conta_id = '" & ContaContabilPagadora & "' " & _
                                     "   AND Titulo = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text
                        End If
                        arrSQL.Add(strSQL)

                        '' rotina de update de contas a pagar inicio
                        If FinanceiroNovo Then
                            strSQL = "UPDATE TITULOS" & vbCrLf & _
                                     "       SET CHEQUE     = 1, " & vbCrLf & _
                                     "           NumeroDoCheque = " & intNumeroCheque & vbCrLf & _
                                     "   WHERE RECPAG = 'P'" & vbCrLf & _
                                     "       AND TITULO_ID = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text & vbCrLf & _
                                     "" & vbCrLf
                        Else
                            strSQL = " UPDATE ContasAPagar" & vbCrLf & _
                                     "           SET Cheque = 'S'," & vbCrLf & _
                                     "               NumeroDoCheque = " & intNumeroCheque & vbCrLf & _
                                     "       WHERE Registro_ID = " & GridConsultaTitulos1.Rows(intPos).Cells(0).Text & vbCrLf
                        End If
                        arrSQL.Add(strSQL)

                        'ATUALIZA NUMERADOR
                        strSQL = " UPDATE BANCOSXCONTAS" & vbCrLf & _
                                 "    SET NumCheque = " & intNumeroCheque & vbCrLf & _
                                 "  WHERE Empresa_Id       = '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                 "    AND EndEmpresa_Id    = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & vbCrLf & _
                                 "    AND Banco_Id         = " & DdlBancoPagador.SelectedValue & vbCrLf & _
                                 "    AND Agencia_Id       = '" & DdlContaPagadora.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                 "    AND DigitoAgencia_Id = '" & DdlContaPagadora.SelectedValue.Split("-")(1) & "'" & vbCrLf & _
                                 "    AND Conta_Id         = '" & DdlContaPagadora.SelectedValue.Split("-")(2) & "'" & vbCrLf & _
                                 "    AND DigitoConta_Id   = '" & DdlContaPagadora.SelectedValue.Split("-")(3) & "'" & vbCrLf
                        arrSQL.Add(strSQL)

                        '' rotina de update de contas a pagar fim 
                        intPosCheque += 1
                        intNumeroCheque += 1
                    End If
                Next

                Arquivo.Close()

                If Banco.GravaBanco(arrSQL) Then
                    Dim strMensagem As String = ""
                    If DdlBancoPagador.SelectedValue = 756 Or DdlBancoPagador.SelectedValue = 1 Then
                        If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "04854422" Then
                            System.IO.File.Copy(Server.MapPath("Files/Cheque.txt"), HttpContext.Current.Session("printerFinanceiro"))
                        Else
                            Dim objImpressao As New cheque.ImpressaoCheques()
                            strMensagem = objImpressao.ImprimirCheques(dblValor, strValorExtenso, strNome, strCidade, dateData, intNumero, Convert.ToInt32(strBanco))
                        End If
                    Else
                        Dim objImpressao As New cheque.ImpressaoCheques()
                        testeUrl = objImpressao.Url
                        strMensagem = objImpressao.ImprimirCheques(dblValor, strValorExtenso, strNome, strCidade, dateData, intNumero, Convert.ToInt32(strBanco))
                    End If
                    MsgBox(Me.Page, "Os cheques foram emitidos com sucesso! " & strMensagem, eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage")) & ". " & testeUrl)
                End If

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "EmissaoDeCheques")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkEmitirPdf_Click(sender As Object, e As EventArgs) Handles lnkEmitirPdf.Click
        ImprimirCheque(False)
    End Sub
End Class