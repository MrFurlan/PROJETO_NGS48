Imports NGS.Lib.Negocio

Public Class CompensacaoDeAdiantamentos
    Inherits BasePage

    Private Sql As String
    Private SqlArray As New ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CompensacaoDeAdiantamentos", "ACESSAR") Then
                CarregarUnidade()
                ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "", False)
                BuscarIndexadores()
                CarregarFinalidadeFinanceiro("P")
                Limpar()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteADTO" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)
            HttpContext.Current.Session("txtCnpjDoCliente") = strCliente(0)
            HttpContext.Current.Session("txtEndDoCliente") = strCliente(1)
            HttpContext.Current.Session("codigoSafra") = ""

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedidoExtrato" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedidoEfetivo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedidoEfetivo.Click
        Try
            Dim parameters As New Dictionary(Of String, Object)
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)

            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("pedidoefetivo") = txtPedidoEfetivo.Text

            Popup.ConsultaDePedidos(Me.Page, "objPedidoExtrato" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub bntAdiantamento_Click(sender As Object, e As EventArgs) Handles bntAdiantamento.Click
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
        ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cliente/Fornecedor não foi selecionado!")
        ElseIf ddlFinalidadeFinanceira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Conta do Adiantamento não foi selecionada!")
        ElseIf txtCodigoCliente.Value.ToString.Length > 0 Then

            lnkBaixar.Parent.Visible = False
            lnkConsultar.Parent.Visible = False

            gridConsultaTitulos.DataSource = Nothing
            gridConsultaTitulos.DataBind()

            Dim ff = New CarteiraFinanceira(ddlFinalidadeFinanceira.SelectedValue)
            Dim Parametros As New Hashtable
            Parametros.Add("Titulo", 0)
            Parametros.Add("Empresa", ddlEmpresa.SelectedValue.Split("-")(0))
            Parametros.Add("EndEmpresa", ddlEmpresa.SelectedValue.Split("-")(1))
            Parametros.Add("Cliente", txtCodigoCliente.Value.Split("-")(0))
            Parametros.Add("EndCliente", txtCodigoCliente.Value.Split("-")(1))

            Parametros.Add("ContaContabil", ff.CodigoContaCliente)
            Parametros.Add("Moeda", ddlMoeda.SelectedValue)
            Parametros.Add("DescMoeda", ddlMoeda.SelectedItem.Text)
            Parametros.Add("Formulario", "Financeiro")

            If txtPedido.Text.Length > 0 Then Parametros.Add("Pedido", txtPedido.Text)

            Session("Parametros" & HID.Value) = Parametros
            ucConsultaAdiantamentos.BindGridView()
            Popup.ConsultaDeAdiantamentos(Me.Page, "Financeiro" & HID.Value)
        End If
    End Sub

    Protected Sub rdReceber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarFinalidadeFinanceiro("R")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdPagar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarFinalidadeFinanceiro("P")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkDataBaixa_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkDataBaixa.Checked Then
                txtDataBaixa.Enabled = True
            Else
                txtDataBaixa.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBaixar_Click(sender As Object, e As EventArgs) Handles lnkBaixar.Click
        Try
            Dim saldoAdiantamento As Decimal = CDec(gridAdiantamentos.Rows(0).Cells(11).Text)

            If txtValorTotal.Value > saldoAdiantamento Then
                MsgBox(Me.Page, "Valor selecionado não pode ser maior que o Saldo do Adiantamento!")
                Exit Sub
            End If

            If chkDataBaixa.Checked Then
                Dim tituloAdiantamento As New Titulo(gridAdiantamentos.Rows(0).Cells(1).Text)

                If CDate(txtDataBaixa.Text) < tituloAdiantamento.Baixa Then
                    MsgBox(Me.Page, "Data da Baixa para Compensação não pode ser menor que o data do Adiantamento!")
                    Exit Sub
                End If
            End If

            Dim Sequencia As Integer

            Sql = "  SELECT ISNULL(MAX(Sequencia_Id), 0) + 1 AS Sequencia  " & vbCrLf &
                   "    FROM vw_AdiantamentosXBaixas" & vbCrLf &
                   "   WHERE TituloAdiantamento =" & gridAdiantamentos.Rows(0).Cells(1).Text & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "AdiantamentosXBaixas").Tables(0).Rows
                Sequencia = Dr("Sequencia")
            Next

            Dim ff = New CarteiraFinanceira(ddlFinalidadeFinanceira.SelectedValue)

            Dim CodigoBancoPagador As Integer = 0
            Dim CodigoAgenciaPagadora As String = String.Empty
            Dim DigitoAgenciaPagadora As String = String.Empty
            Dim ContaPagadora As String = String.Empty
            Dim DigitoContaPagadora As String = String.Empty
            Dim TipoDaContaPagadora As String = String.Empty
            Dim ContaContabilPagadora As String = String.Empty

            Sql = " SELECT bxc.Banco_Id, bxc.Agencia_Id, bxc.DigitoAgencia_Id, bxc.Conta_Id, bxc.DigitoConta_Id, isnull(bxc.TipoConta,'C') as TipoConta, bxc.ContaContabil, bxc.Observacoes, Pc.Titulo" & vbCrLf &
                  "   FROM BancosXContas bxc" & vbCrLf &
                  "  Inner Join PlanoDeContas PC" & vbCrLf &
                  "     on Bxc.ContaContabil = Pc.Conta_Id" & vbCrLf &
                  "  WHERE bxc.Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                  "    AND bxc.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                  "    AND bxc.ContaContabil = '" & ff.CodigoContaCliente & "'" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                CodigoBancoPagador = Dr("Banco_Id")
                CodigoAgenciaPagadora = Dr("Agencia_Id")
                DigitoAgenciaPagadora = Dr("DigitoAgencia_Id")
                ContaPagadora = Dr("Conta_Id")
                DigitoContaPagadora = Dr("DigitoConta_Id")
                TipoDaContaPagadora = Dr("TipoConta")
                ContaContabilPagadora = Dr("ContaContabil")
            Next

            If CodigoBancoPagador = 0 Then
                MsgBox(Me.Page, "Verificar Bancos X Contas, não foi encontrado a conta de Compensação de Adiantamento!")
                Exit Sub
            End If

            SqlArray.Clear()

            For Each row As GridViewRow In gridConsultaTitulos.Rows
                Dim chkTitulo As CheckBox = CType(row.FindControl("chkGridTitulos"), CheckBox)
                If (chkTitulo.Checked) Then
                    Dim ValorBxAdtoOficial As Decimal = CDec(row.Cells(6).Text)
                    Dim ValorBxAdtoMoeda As Decimal = CDec(row.Cells(5).Text)
                    Dim ValorVariacao As Decimal = 0

                    Dim tit As New Titulo(row.Cells(1).Text)

                    If tit.ValorDoDocumento > 0 Then

                        tit.IUD = "U"
                        tit.CodigoProvisao = 1

                        tit.CodigoEmpresaPagadora = ddlEmpresa.SelectedValue.Split("-")(0)
                        tit.EndEmpresaPagadora = ddlEmpresa.SelectedValue.Split("-")(1)

                        If chkDataBaixa.Checked Then
                            If CDate(txtDataBaixa.Text) < tit.Prorrogacao Then
                                MsgBox(Me.Page, "Data da Baixa para Compensação não pode ser menor que o data de Vencimento do Título!")
                                Exit Sub
                            End If

                            tit.Baixa = CDate(txtDataBaixa.Text)
                            tit.Prorrogacao = CDate(txtDataBaixa.Text)
                        Else
                            tit.Baixa = tit.Prorrogacao
                        End If

                        If Not Funcoes.VerificaAcesso(tit.CodigoEmpresa, tit.EnderecoEmpresa, tit.Baixa.ToString("dd/MM/yyyy"), "CONTABIL") Then
                            MsgBox(Me.Page, "Movimento Contábil do dia " & tit.Baixa.ToString("dd/MM/yyyy") & " já Fechado para esta data...")
                            Exit Sub
                        End If

                        If Not Funcoes.VerificaAcesso(tit.CodigoEmpresaPagadora, tit.EndEmpresaPagadora, tit.Baixa.ToString("dd/MM/yyyy"), "CONTABIL") Then
                            MsgBox(Me.Page, "Movimento Contábil do dia " & tit.Baixa.ToString("dd/MM/yyyy") & " da Empresa Pagadora já Fechado para esta data...")
                            Exit Sub
                        End If

                        tit.CodigoBancoPagador = CodigoBancoPagador
                        tit.CodigoAgenciaPagadora = CodigoAgenciaPagadora
                        tit.DigitoAgenciaPagadora = DigitoAgenciaPagadora
                        tit.ContaPagadora = ContaPagadora
                        tit.DigitoContaPagadora = DigitoContaPagadora
                        tit.TipoContaPagadora = TipoDaContaPagadora
                        tit.ContaContabilPagadora = ContaContabilPagadora
                        tit.UsuarioBaixa = Session("ssNomeUsuario")
                        tit.UsuarioBaixaData = CDate(Today).ToString("yyyy/MM/dd")

                        Dim obs As String = tit.ObservacoesControleInterno
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Baixado pelo Processo CompensacaoDeAdiantamentos em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        Else
                            obs = "Baixado pelo Processo CompensacaoDeAdiantamentos em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        End If

                        tit.ObservacoesControleInterno = obs

                        tit.SalvarSql(SqlArray)

                        Sql = " INSERT INTO AdiantamentosXBaixas" & vbCrLf &
                                " (Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, Sequencia_Id, Titulo, ValorOficial, ValorMoeda, VariacaoOficial, DataBaixa)" & vbCrLf &
                                " VALUES('" & ddlEmpresa.SelectedValue.Split("-")(0) & "'," & ddlEmpresa.SelectedValue.Split("-")(1) & ", '" & txtCodigoCliente.Value.Split("-")(0) & "', " & txtCodigoCliente.Value.Split("-")(1) & "," & vbCrLf &
                                            gridAdiantamentos.Rows(0).Cells(0).Text & ", " & Sequencia & ", " & row.Cells(1).Text & ", " & Str(ValorBxAdtoOficial) & ", " & Str(ValorBxAdtoMoeda) & "," & Str(ValorVariacao) & ",'" & CDate(row.Cells(2).Text).ToSqlDate & "')"

                        SqlArray.Add(Sql)

                        Sequencia += 1
                    End If

                End If
            Next

            If SqlArray.Count > 0 Then
                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "Compensações realizadas com sucesso.")
                    Limpar()
                Else
                    MsgBox(Me.Page, Session("ssMessage"))
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada!")
            ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
                MsgBox(Me.Page, "Cliente/Fornecedor não foi selecionado!")
            ElseIf ddlFinalidadeFinanceira.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Conta do Adiantamento não foi selecionada!")
            ElseIf Not ddlMoeda.SelectedValue = 1 AndAlso ddlIndexador.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Indexador não foi selecionado!")
            Else
                ConsultarTitulos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível limpar.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CompensacaoDeAdiantamentos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteADTO" & HID.Value) Is Nothing Then
            txtCodigoCliente.Value = CType(Session("objClienteADTO" & HID.Value), [Lib].Negocio.Cliente).Codigo & "-" & CType(Session("objClienteADTO" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(CType(Session("objClienteADTO" & HID.Value), [Lib].Negocio.Cliente).Codigo) & " - " & CType(Session("objClienteADTO" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteADTO" & HID.Value)
        ElseIf Session("objPedidoExtrato" & HID.Value) IsNot Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objPedidoExtrato" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = p.Codigo
            If p.PedidoEfetivo.Length > 0 Then txtPedidoEfetivo.Text = p.PedidoEfetivo

            Session.Remove("objPedidoExtrato" & HID.Value)

            If p.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                rdPagar.Checked = True
                CarregarFinalidadeFinanceiro("P")
            Else
                rdReceber.Checked = True
                CarregarFinalidadeFinanceiro("R")
            End If

        End If
    End Sub

    Private Sub BuscarIndexadores()
        Dim objIndexadores As New [Lib].Negocio.Indexadores()

        If objIndexadores.Selecionar() Then
            For Each objIndexador As [Lib].Negocio.Indexador In objIndexadores
                ddlIndexador.Items.Add(New ListItem(objIndexador.Codigo.ToString() & "-" & objIndexador.Descricao,
                                                    objIndexador.Codigo.ToString()))
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddlIndexador)
    End Sub

    Private Sub CarregarFinalidadeFinanceiro(ByVal Tipo As String)
        ddlFinalidadeFinanceira.Items.Clear()

        Sql = "SELECT Cart.Produto_Id AS Codigo," & vbCrLf &
              "       case" & vbCrLf &
              "			when isnull(Cart.ContaClientes,'') = ''" & vbCrLf &
              "			   then Cart.Produto_Id + '  -  ' + Cart.Descricao" & vbCrLf &
              "			   else Cart.Produto_Id + '  -  ' + Cart.Descricao + ' (' + Cart.ContaClientes + '-' + pl.Titulo + ')'" & vbCrLf &
              "		  end AS Descricao" & vbCrLf &
              "" & vbCrLf &
              "  FROM ComprasXProdutos Cart" & vbCrLf &
              "  LEFT JOIN PlanoDeContas pl " & vbCrLf &
              "         on pl.Conta_Id = Cart.ContaClientes " & vbCrLf

        Sql &= " Where Cart.Adiantamento = 'S' AND  Cart.Classificacao = '" & Tipo & "'" & vbCrLf &
                " Order By Produto_Id" & vbCrLf

        ddlFinalidadeFinanceira.DataValueField = "Codigo"
        ddlFinalidadeFinanceira.DataTextField = "Descricao"
        ddlFinalidadeFinanceira.DataSource = Banco.ConsultaDataSet(Sql, "Carteiras")
        ddlFinalidadeFinanceira.DataBind()

        ddlFinalidadeFinanceira.Items.Insert(0, "")
        ddlFinalidadeFinanceira.SelectedIndex = 0
    End Sub

    Public Sub CarregarAdiantamento(ByVal dsAdiantamento As DataSet, ByVal index As Integer)

        divAdiantamentos.Visible = True

        Dim dsAdto As DataTable = dsAdiantamento.Tables(0).Clone()
        dsAdto.ImportRow(dsAdiantamento.Tables(0).Rows(index))

        gridAdiantamentos.DataSource = dsAdto
        gridAdiantamentos.DataBind()

        lnkConsultar.Parent.Visible = True

        Dim titulo As New Titulo(dsAdiantamento.Tables(0).Rows(index).Item(1))

        If titulo.ValorDoDocumento > 0 Then
            hidDataBaixaAdiantamento.Value = titulo.Baixa
        End If

        If CInt(dsAdiantamento.Tables(0).Rows(index).Item(2)) > 0 Then
            txtPedido.Enabled = False
            txtPedidoEfetivo.Enabled = False
        End If

    End Sub

    Private Sub ConsultarTitulos()

        Dim Tipo As String = IIf(rdPagar.Checked, "ContasAPagar", "ContasAReceber")
        Dim Valor As Decimal = 0

        Sql = "  SELECT CP.Registro_Id AS Registro, " & vbCrLf &
              "         convert(varchar(10),CP.Prorrogacao,103) AS Vencimento, " & vbCrLf &
              "         Cli.Nome AS Cliente, CP.Historico, isnull(CP.MoedaValorDoDocumento, 0) AS Dolar, CP.ValorDoDocumento AS Valor, " & vbCrLf &
              "         ISNULL(CP.MoedaValorLiquido, 0) AS MoedaLiquido, CP.ValorLiquido AS ValorLiquido, " & vbCrLf &
              "         UsuarioLiberacao as Liberado, CP.Pedido as Pedido, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             WHEN CP.Moeda = 0 THEN 'R$-' + convert(varchar,CP.Moeda) " & vbCrLf &
              "             ELSE " & vbCrLf &
              "                 CASE " & vbCrLf &
              "                     WHEN CP.Moeda = 1  THEN 'R$-' + convert(varchar,CP.Moeda) " & vbCrLf &
              "                     ELSE 'U$-' + convert(varchar,CP.Moeda) " & vbCrLf &
              "                 END " & vbCrLf &
              "         END as Moeda, CP.Indexador, isnull(CP.Grupado,'N') as Grupado, CP.Provisao, CP.Situacao " & vbCrLf &
              "    FROM " & Tipo & " CP " & vbCrLf &
              "    INNER JOIN NotaFiscalXTitulo NFXT" & vbCrLf &
              "      ON CP.Empresa     = NFXT.Empresa_Id" & vbCrLf &
              "     AND CP.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf &
              "     AND CP.Registro_Id = NFXT.Titulo_Id" & vbCrLf &
              "    INNER JOIN NotasFiscais NF" & vbCrLf &
              "      ON NF.Empresa_Id       = NFXT.Empresa_Id" & vbCrLf &
              "     AND NF.EndEmpresa_Id    = NFXT.EndEmpresa_Id" & vbCrLf &
              "     AND NF.Cliente_Id       = NFXT.Cliente_Id" & vbCrLf &
              "     AND NF.EndCliente_Id    = NFXT.EndCliente_Id" & vbCrLf &
              "     AND NF.EntradaSaida_id  = NFXT.EntradaSaida_id" & vbCrLf &
              "     AND NF.Serie_Id         = NFXT.Serie_Id" & vbCrLf &
              "     AND NF.Nota_Id          = NFXT.Nota_Id" & vbCrLf &
              "   INNER JOIN Clientes Cli" & vbCrLf &
              "      ON CP.Cliente = Cli.Cliente_Id " & vbCrLf &
              "     AND CP.EndCliente = Cli.Endereco_Id" & vbCrLf &
              " WHERE CP.UnidadeDeNegocio = '" & ddlUnidade.SelectedValue & "' " & vbCrLf &
              "   AND CP.Empresa          = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
              "   AND CP.EndEmpresa       = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
              "   AND CP.Cliente          = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
              "   AND CP.EndCliente       = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf &
              "   AND CP.Provisao    = 2 " & vbCrLf &
              "   AND CP.Situacao    = 1 " & vbCrLf &
              "   AND CP.Grupado     = 'N'" & vbCrLf &
              "   AND CP.Moeda       = " & ddlMoeda.SelectedValue & vbCrLf &
              "   AND NF.Movimento  >= " & CDate(gridAdiantamentos.Rows(0).Cells(5).Text).ToString("yyyy-MM-dd") & vbCrLf


        If ddlIndexador.SelectedIndex > 0 Then
            Sql &= "   AND CP.Indexador   = " & ddlIndexador.SelectedValue & vbCrLf
        End If

        If txtPedido.Text.Length > 0 AndAlso CInt(txtPedido.Text) > 0 Then
            Sql &= "   AND CP.Pedido      = " & txtPedido.Text & vbCrLf
        End If

        Sql &= " ORDER BY CP.Prorrogacao, Cli.Nome"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contas")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum Registro encontrado...")
        Else
            lnkConsultar.Parent.Visible = True
            lnkBaixar.Parent.Visible = True

            gridConsultaTitulos.DataSource = ds
            gridConsultaTitulos.DataBind()

            Dim i As Integer = 0
            While i < gridConsultaTitulos.Rows.Count
                Dim strMoeda() As String = gridConsultaTitulos.Rows(i).Cells(8).Text.ToString.Split("-")
                If strMoeda(0) = "U$" Then
                    gridConsultaTitulos.Rows(i).ForeColor = System.Drawing.Color.Red
                End If

                If (gridConsultaTitulos.Rows(i).Cells(8).Text.Equals("R$-1")) Then
                    Valor = Valor + CDec(gridConsultaTitulos.Rows(i).Cells(6).Text)
                Else
                    Valor = Valor + CDec(gridConsultaTitulos.Rows(i).Cells(5).Text)
                End If

                i += 1
            End While

            If Valor > 0 Then
                lblTotalRegistroAgrupado.Parent.Visible = True
                lblTotalRegistroAgrupado.Text = " Total de Título(s) no valor de: " & String.Format("{0:N2}", Valor)
            End If

            rowDolar.Visible = True

        End If

    End Sub

    Private Sub Limpar()

        Session.Remove("objClienteADTO" & HID.Value)
        Session.Remove("objPedidoExtrato" & HID.Value)

        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty

        txtPedido.Enabled = True
        txtPedidoEfetivo.Enabled = True

        txtPedido.Text = String.Empty
        txtPedidoEfetivo.Text = String.Empty

        chkDataBaixa.Checked = False
        txtDataBaixa.Text = Format(DateTime.Now, "dd/MM/yyyy")
        txtDataBaixa.Enabled = False

        lblTotalRegistroAgrupado.Text = String.Empty
        lblTotalRegistroAgrupado.Parent.Visible = False

        rowDolar.Visible = False
        txtRealDolar.Value = String.Empty
        HiddenIndexador.Value = String.Empty
        txtValorTotal.Value = String.Empty
        hidDataBaixaAdiantamento.Value = String.Empty

        lnkBaixar.Parent.Visible = False
        lnkConsultar.Parent.Visible = False

        gridConsultaTitulos.DataSource = Nothing
        gridConsultaTitulos.DataBind()

        gridAdiantamentos.DataSource = Nothing
        gridAdiantamentos.DataBind()

        divAdiantamentos.Visible = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaAdiantamentos.SetarHID(HID.Value)
    End Sub

    Protected Sub chkAllTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If gridConsultaTitulos.Rows.Count > 0 Then
                Dim chkAllTitulos As CheckBox = CType(sender, CheckBox)

                Dim passed As Boolean = False
                For Each rowgrid As GridViewRow In gridConsultaTitulos.Rows
                    Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("ChkGridTitulos"), CheckBox)

                    If Not chkTitulo.Enabled Then
                        chkAllTitulos.Checked = False
                        Exit Sub
                    End If

                    If chkAllTitulos.Checked Then

                        If hidDataBaixaAdiantamento.Value > CDate(rowgrid.Cells(2).Text) Then
                            chkAllTitulos.Checked = False
                            chkTitulo.Checked = False
                            HiddenIndexador.Value = String.Empty
                            txtRealDolar.Value = String.Empty
                            MsgBox(Me.Page, "Data do título a compensar não pode ser menor que a data do Adiantamento.")
                            Exit Sub
                        End If

                        Dim strMoeda As String = rowgrid.Cells(8).Text.ToString.Split("-")(1)

                        If Not passed Then
                            txtRealDolar.Value = strMoeda
                            HiddenIndexador.Value = rowgrid.Cells(9).Text
                            passed = True
                        End If
                        chkTitulo.Checked = IIf(strMoeda <> txtRealDolar.Value OrElse rowgrid.Cells(9).Text <> HiddenIndexador.Value, False, True)
                    Else
                        chkTitulo.Checked = False
                        HiddenIndexador.Value = String.Empty
                        txtRealDolar.Value = String.Empty
                    End If
                Next
                TotalizadorTitulosAgrupados()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTitulo As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkTitulo.NamingContainer, GridViewRow)

        If chkTitulo.Checked Then

            Dim Titulo As New Titulo(CInt(gridConsultaTitulos.Rows(row.RowIndex).Cells(1).Text()))

            If hidDataBaixaAdiantamento.Value > CDate(gridConsultaTitulos.Rows(row.RowIndex).Cells(2).Text) Then
                chkTitulo.Checked = False
                HiddenIndexador.Value = String.Empty
                txtRealDolar.Value = String.Empty
                MsgBox(Me.Page, "Data do título a compensar não pode ser menor que a data do Adiantamento.")
                Exit Sub
            End If

            If Titulo.CodigoProvisao = eProvisao.Provisao Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Titulo em Provisao não pode ser usado no Agrupamento.")
                Exit Sub
            End If

            Dim strMoeda() As String = gridConsultaTitulos.Rows(row.RowIndex).Cells(8).Text.ToString.Split("-")
            If txtRealDolar.Value.ToString.Length = 0 Then
                txtRealDolar.Value = strMoeda(1)
            End If

            If HiddenIndexador.Value.ToString.Length = 0 Then
                HiddenIndexador.Value = gridConsultaTitulos.Rows(row.RowIndex).Cells(9).Text
            End If

            If Not strMoeda(1) = txtRealDolar.Value Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Agrupamento permitido apenas para Títulos com a mesma moeda.")
                Exit Sub
            End If

            If Not gridConsultaTitulos.Rows(row.RowIndex).Cells(9).Text = HiddenIndexador.Value Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Agrupamento permitido apenas para Títulos com o mesmo indexador.")
                Exit Sub
            End If
        End If
        TotalizadorTitulosAgrupados()
    End Sub

    Public Sub TotalizadorTitulosAgrupados()
        Dim Quantidade As Integer = 0
        Dim Valor As Decimal = 0
        For Each row As GridViewRow In gridConsultaTitulos.Rows
            Dim chkTitulo As CheckBox = CType(row.FindControl("chkGridTitulos"), CheckBox)
            If (chkTitulo.Checked) Then
                Quantidade = Quantidade + 1
                If (row.Cells(8).Text.Equals("R$-1")) Then
                    Valor = Valor + CDec(row.Cells(6).Text)
                Else
                    Valor = Valor + CDec(row.Cells(5).Text)
                End If
            End If
        Next
        If Quantidade > 0 Then
            lnkBaixar.Parent.Visible = True
        Else
            lnkBaixar.Parent.Visible = False
        End If
        If Quantidade = 0 Then
            lblTotalRegistroAgrupado.Text = String.Empty
            lblTotalRegistroAgrupado.Parent.Visible = False
            Dim ckAll As CheckBox = CType(gridConsultaTitulos.HeaderRow.FindControl("chkAllTitulos"), CheckBox)
            ckAll.Checked = False
            txtValorTotal.Value = 0
        Else
            lblTotalRegistroAgrupado.Text = Quantidade & " Título(s) a selecionado(s) no valor total de: " & String.Format("{0:N2}", Valor)
            lblTotalRegistroAgrupado.Parent.Visible = True
            txtValorTotal.Value = Valor
        End If
    End Sub

End Class