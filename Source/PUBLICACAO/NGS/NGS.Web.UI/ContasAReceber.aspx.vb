Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class ContasAReceber
    Inherits BasePage

#Region "Atributos / Propriedades"

    Dim Sql As String
    Dim Sqla As String

    Dim SqlArray As New ArrayList
    Dim Unidades As New ArrayList

    Dim DS As DataSet

    Dim Codigo As String
    Dim Descricao As String
    Dim Cliente As String

    Dim campo() As String
    Dim Valor As Decimal
    Dim Data As String

    Dim Registro As Integer
    Dim TemRegistro As String

    Dim Raz_Empresa As String
    Dim Raz_EndEmpresa As Integer
    Dim Raz_Conta As String
    Dim Raz_Cliente As String
    Dim Raz_EndCliente As String
    Dim Raz_UnidadeDeNegocio As String
    Dim Raz_ValorOficial As String
    Dim Raz_ValorMoeda As String

    Dim Raz_Historico As String
    Dim Raz_DebitoCredito As String

    Dim ContaClientes As String
    Dim ContaDescontos As String
    Dim ContaDeducoes As String
    Dim ContaJuros As String
    Dim ContaAcrescimos As String

    Dim MoedaJuros As String
    Dim MoedaAcrescimos As String
    Dim MoedaDescontos As String
    Dim MoedaDeducoes As String

    Dim MensagemParcial As String
    Dim strJavaScript As String
    Private objPedido As [Lib].Negocio.Pedido

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not Page.IsPostBack AndAlso IsConnect Then
            'If Funcoes.VerificaPermissao("ContasAReceber", "ACESSAR") Then
            '    If Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
            '        chkPrevisao.Visible = True
            '        chkProvisao.Visible = True
            '    End If
            '    CargaUnidadeDeNegocioEmpresaCliente()
            '    TiposDePagamentos()
            '    Provisoes()
            '    CarteiraDoTitulo()
            '    Carteiras()
            '    BuscarMoedas()
            '    BuscarIndexadores()
            '    Limpar(True)
            '    LimparConsultaTitulos(True)
            '    txtMovimento.Text = CDate(Today).ToString("dd/MM/yyyy")
            '    hdnMovimentoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")
            '    txtPeriodoInicialConsultaTitulos.Text = CDate(Today).ToString("dd/MM/yyyy")
            '    txtPeriodoFinalConsultaTitulos.Text = CDate(Today).ToString("dd/MM/yyyy")
            '    TabContainer1.ActiveTabIndex = 0
            '    VerificaUnidade()
            '    ddl.Carregar(ddlSelecionarHist, CarregarDDL.Tabela.Historico)
            'Else
            MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!")
            Exit Sub
            'End If
        Else
            'If DdlBancoRecebedor IsNot Nothing Then
            '    For Each li As ListItem In DdlBancoRecebedor.Items
            '        li.Attributes("title") = li.Text
            '    Next
            'End If
        End If
    End Sub

    Public Sub MostrarCotacao()
        If Not String.IsNullOrWhiteSpace(txtProrrogacao.Text) AndAlso IsDate(txtProrrogacao.Text) AndAlso Not String.IsNullOrWhiteSpace(ddlIndexador.SelectedValue) Then
            lblCotacao.Text = Funcoes.PegarValorConversao(ddlIndexador.SelectedValue, txtProrrogacao.Text).ToString("N4")
            lblDescCotacao.Text = ddlIndexador.SelectedItem.Text.Split("-")(1)
        End If
    End Sub

    Protected Sub ddlIndexador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlIndexador.SelectedIndexChanged
        Try
            MostrarCotacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlSelecionarHist_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSelecionarHist.SelectedIndexChanged
        Try
            txtHistorico.Text = ddlSelecionarHist.SelectedItem.Text
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub TotalizadorTitulosAgrupados()
        Dim Quantidade As Integer = 0
        Dim Valor As Decimal = 0
        For Each row As GridViewRow In GridConsultaTitulos.Rows
            Dim chkTitulo As CheckBox = CType(row.FindControl("ChkGridTitulos"), CheckBox)
            If (chkTitulo.Checked) Then
                Quantidade = Quantidade + 1
                If (row.Cells(8).Text.Equals("R$-1")) Then
                    Valor = Valor + CDec(row.Cells(7).Text)
                Else
                    Valor = Valor + CDec(row.Cells(6).Text)
                End If
            End If
        Next

        If Quantidade > 1 Then
            lnkAgrupar.Parent.Visible = True
        Else
            lnkAgrupar.Parent.Visible = False
        End If

        If Quantidade > 0 Then
            lblTotalRegistroAgrupado.Parent.Visible = True
            lblTotalRegistroAgrupado.Text = Quantidade & " Títulos a selecionados no valor total de: " & String.Format("{0:N2}", Valor)
        Else
            lblTotalRegistroAgrupado.Text = String.Empty
            lblTotalRegistroAgrupado.Parent.Visible = False
            Dim ckAll As CheckBox = CType(GridConsultaTitulos.HeaderRow.FindControl("chkAllTitulos"), CheckBox)
            ckAll.Checked = False
        End If
    End Sub

    Private Sub BuscarMoedas()
        Dim objMoedas As New [Lib].Negocio.Moedas()

        ddlMoeda.Items.Clear()
        ddlMoeda.Items.Add(New ListItem("", 0))
        For Each objMoeda As [Lib].Negocio.Moeda In objMoedas
            ddlMoeda.Items.Add(New ListItem(objMoeda.Codigo.ToString() & "-" & objMoeda.Descricao, objMoeda.Codigo.ToString()))
        Next
    End Sub

    Private Sub BuscarIndexadores()
        Dim objIndexadores As New [Lib].Negocio.Indexadores()

        If objIndexadores.Selecionar() Then
            For Each objIndexador As [Lib].Negocio.Indexador In objIndexadores
                ddlIndexador.Items.Add(New ListItem(objIndexador.Codigo.ToString() & "-" & objIndexador.Descricao, _
                                                    objIndexador.Codigo.ToString()))
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddlIndexador)
    End Sub

    Sub VerificaUnidade()
        Sql = " SELECT ISNULL(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        ISNULL(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        ISNULL(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              "   FROM Usuarios" & vbCrLf & _
              "  WHERE Usuario_Id = '" & Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidadeDeNegocioEmpresaCliente.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresaCliente()
            DdlEmpresaCliente.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

#Region "Consulta Titulos"

    Protected Sub btnBuscaPedidoConsultaTitulos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaPedidoConsultaTitulos.Click
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Or String.IsNullOrWhiteSpace(txtCodigoClienteConsTitulo.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoClienteConsTitulo.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)
            HttpContext.Current.Session("txtCnpjDoCliente") = strCliente(0)
            HttpContext.Current.Session("txtEndDoCliente") = strCliente(1)

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objContasAReceber" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarTitulos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultarTitulos.Click
        TitulosConsulta()
    End Sub

    Sub TitulosConsulta()
        Dim dra As DataRow
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue

        Sql = " SELECT CR.Registro_Id AS Registro, convert(varchar(10),CR.Prorrogacao,103) as Vencimento, " & vbCrLf & _
              "        Clientes.Nome AS Cliente, Historico, isnull(CR.MoedaValorLiquido, 0) AS Dolar, CR.ValorLiquido AS Valor, CR.Pedido as Pedido," & vbCrLf & _
              "        CASE " & vbCrLf & _
              "            WHEN CR.Moeda = 0 then 'R$-' + convert(varchar,CR.Moeda) " & vbCrLf & _
              "            ELSE " & vbCrLf & _
              "                CASE " & vbCrLf & _
              "                    WHEN CR.Moeda = 1  then 'R$-' + convert(varchar,CR.Moeda) " & vbCrLf & _
              "                    ELSE 'U$-' + convert(varchar,CR.Moeda) " & vbCrLf & _
              "                END " & vbCrLf & _
              "        END as Moeda, CR.Indexador, isnull(CR.grupado,'N') as Grupado, CR.Provisao " & vbCrLf & _
              "   FROM ContasAReceber CR" & vbCrLf & _
              "   LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf & _
              "     ON CR.Empresa     = NFXT.Empresa_Id" & vbCrLf & _
              "    AND CR.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf & _
              "    AND CR.Registro_Id = NFXT.Titulo_Id" & vbCrLf & _
              "   LEFT JOIN FaturaDeFreteXTitulo FFXT" & vbCrLf & _
              "     ON CR.Registro_Id = FFXT.Titulo_Id" & vbCrLf & _
              "   LEFT JOIN FaturasDeFretesXItens FFXI" & vbCrLf & _
              "     ON FFXI.EmpresaPagadora_Id    = FFXT.Empresa_Id" & vbCrLf & _
              "    AND FFXI.EndEmpresaPagadora_Id = FFXT.EndEmpresa_Id" & vbCrLf & _
              "    AND FFXI.Fatura_Id             = FFXT.Fatura_Id" & vbCrLf & _
              "  INNER JOIN Clientes ON CR.Cliente = Clientes.Cliente_Id AND CR.EndCliente = Clientes.Endereco_Id" & vbCrLf

        If RbCancelado.Checked Then
            Sql &= "  WHERE CR.Situacao <> 1 " & vbCrLf
        Else
            Sql &= "  WHERE CR.Situacao = 1 " & vbCrLf
        End If

        If RbAtivo.Checked Then
            If chkPrevisao.Visible = True And chkProvisao.Visible = True Then
                If chkPrevisao.Checked = True AndAlso chkProvisao.Checked = True Then
                    Sql &= "    AND (Provisao = 2 OR Provisao = 3) " & vbCrLf
                ElseIf chkPrevisao.Checked = True Then
                    Sql &= "    AND Provisao = 2 " & vbCrLf
                ElseIf chkProvisao.Checked = True Then
                    Sql &= "    AND Provisao = 3 " & vbCrLf
                Else
                    Sql &= "    AND Provisao = 2 " & vbCrLf
                End If
            Else
                Sql &= "    AND Provisao = 2 " & vbCrLf
            End If
        End If

        If RbBaixado.Checked Then
            Sql &= "    AND Provisao = 1 " & vbCrLf
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            Sql &= "    AND CR.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            Sql &= "    AND CR.Empresa = '" & Campo(0) & "'" & vbCrLf   'Empresa
            Sql &= "    AND CR.EndEmpresa = " & Campo(1) & vbCrLf    'Endereco da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= "    AND (ISNULL(NFXT.Nota_Id,0) = " & txtNumNota.Text & " OR ISNULL(FFXI.Nota_Id,0) = " & txtNumNota.Text & ")"
        Else
            Cliente = txtCodigoClienteConsTitulo.Value
            Campo = Cliente.Split("-")
            If Campo(0) <> "" Then
                Sql &= "    AND CR.Cliente = '" & Campo(0) & "'" & vbCrLf & _
                       "    AND CR.EndCliente = " & Campo(1) & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
                Sql &= "    AND CR.Pedido= '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
                Sql &= "    AND CR.Pedido= '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
            End If

            If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
                Sql &= "    AND CR.Prorrogacao between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
            End If
        End If

        Sql &= "  ORDER BY CR.Prorrogacao, Clientes.Nome, CR.Registro_Id"

        DS = Banco.ConsultaDataSet(Sql, "Contas")

        If DS Is Nothing Then
            GridConsultaTitulos.DataBind()
            MsgBox(Me.Page, "Nenhum Registro encontrado...")
        ElseIf DS.Tables(0).Rows.Count = 0 Then
            GridConsultaTitulos.DataBind()
            MsgBox(Me.Page, "Nenhum Registro encontrado...")
        Else
            For Each dra In DS.Tables(0).Rows
                If dra("Moeda") = "3" Then
                    dra("Cliente") = "$" & dra("Cliente")
                End If
            Next

            GridConsultaTitulos.DataSource = DS
            GridConsultaTitulos.DataBind()

            While i < GridConsultaTitulos.Rows.Count
                Dim strMoeda() As String = GridConsultaTitulos.Rows(i).Cells(8).Text.ToString.Split("-")
                If strMoeda(0) = "U$" Then
                    GridConsultaTitulos.Rows(i).ForeColor = Drawing.Color.Red
                End If

                If GridConsultaTitulos.Rows(i).Cells(10).Text = "S" Or GridConsultaTitulos.Rows(i).Cells(10).Text = "M" Then 'GRUPADO BLOQUEIA
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                End If

                If DS.Tables(0).Rows(i).Item("Provisao") = "1" Then
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                    GridConsultaTitulos.Rows(i).Cells(2).ForeColor = Drawing.Color.Red
                    GridConsultaTitulos.Rows(i).Cells(2).ToolTip = "BAIXADO"
                End If

                If RbCancelado.Checked Then
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                End If

                i += 1
            End While

            If Not RbCancelado.Checked Then
                lnkBaixar.Parent.Visible = True
                lnkEmiteReciboGeral.Parent.Visible = True
                lnkReprogramar.Parent.Visible = True
            End If
        End If
    End Sub

#End Region

#Region "Manutencao dos Titulos"

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        ddl.Carregar(DdlUnidadeDeNegocioEmpresaCliente, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(DdlEmpresaPagadora, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

#End Region

    Protected Sub DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
    End Sub

    Sub CargaEmpresaCliente()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaCliente.Items.Clear()

        Sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              "  FROM GruposXEmpresas " & vbCrLf & _
              " INNER JOIN Clientes" & vbCrLf & _
              "    ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaCliente.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlEmpresaCliente.Items.Insert(0, "")
        DdlEmpresaCliente.SelectedIndex = 0
    End Sub

    Sub Empresas()

        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Unidade As String = ""

        Dim Cnpj As String

        Sql = " SELECT GxE.Empresa_Id as Unidade" & vbCrLf & _
              "   FROM Clientes Cli" & vbCrLf & _
              "  INNER JOIN GruposXEmpresas GxE" & vbCrLf & _
              "     ON Cli.Cliente_Id = GxE.Cliente_Id " & vbCrLf & _
              "    AND Cli.Endereco_Id = GxE.EndCliente_Id" & vbCrLf & _
              "  WHERE GxE.Cliente_Id = '" & Session("ssEmpresa") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Unidades.Add(Dr("Unidade"))
        Next

        DdlEmpresaPagadora.Items.Clear()

        Sql = "SELECT Distinct Cli.Cliente_Id AS Codigo, Cli.Endereco_Id, Cli.Reduzido, Cli.Nome, Cli.Cidade, Cli.Estado" & vbCrLf & _
              "  FROM Clientes Cli" & vbCrLf & _
              " INNER JOIN GruposXEmpresas GxE" & vbCrLf & _
              "    ON Cli.Cliente_Id = GxE.Cliente_Id " & vbCrLf & _
              "   AND Cli.Endereco_Id = GxE.EndCliente_Id" & vbCrLf & _
              " WHERE " & vbCrLf

        For index = 0 To Unidades.Count - 1
            If index = 0 Then
                Sql &= " GruposXEmpresas.Empresa_Id = '" & Unidades(index) & "'"
            Else
                Sql &= " or GruposXEmpresas.Empresa_Id = '" & Unidades(index) & "'"
            End If
        Next

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaPagadora.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresaPagadora.Items.Insert(0, "")
        DdlEmpresaPagadora.SelectedIndex = 0

    End Sub

    Sub TiposDePagamentos()

        Sql = "SELECT TipoDePagamento_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(TipoDePagamento_Id AS varchar))) + CAST(TipoDePagamento_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM TiposDePagamentos Order By TipoDePagamento_Id"

        DdlTiposDePagamentos.DataValueField = "Codigo"
        DdlTiposDePagamentos.DataTextField = "Descricao"
        DdlTiposDePagamentos.DataSource = Banco.ConsultaDataSet(Sql, "TiposDePagamentos")
        DdlTiposDePagamentos.DataBind()

        DdlTiposDePagamentos.Items.Insert(0, "")
        DdlTiposDePagamentos.SelectedIndex = 0

    End Sub

    Sub Provisoes()
        Sql = "SELECT Provisao_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(Provisao_Id AS varchar))) + CAST(Provisao_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM Provisoes Order By Provisao_Id"

        DdlProvisoes.DataValueField = "Codigo"
        DdlProvisoes.DataTextField = "Descricao"
        DdlProvisoes.DataSource = Banco.ConsultaDataSet(Sql, "Provisoes")
        DdlProvisoes.DataBind()

        DdlProvisoes.Items.Insert(0, "")
        DdlProvisoes.SelectedIndex = 0

    End Sub

    Private Sub CarteiraDoTitulo()
        Dim objCarteiraDoTitulo As New [Lib].Negocio.ListCarteiraDoTitulo()

        ddlCarteiraDoTitulo.DataValueField = "Codigo"
        ddlCarteiraDoTitulo.DataTextField = "Descricao"
        ddlCarteiraDoTitulo.DataSource = objCarteiraDoTitulo.ToArray()
        ddlCarteiraDoTitulo.DataBind()
    End Sub

    Sub Carteiras()
        Sql = "SELECT  Produto_Id AS Codigo, Produto_Id + '  -  ' + Descricao AS Descricao FROM ComprasXProdutos Where Classificacao = 'R' Order By Produto_Id"

        ddlCarteiras.DataValueField = "Codigo"
        ddlCarteiras.DataTextField = "Descricao"
        ddlCarteiras.DataSource = Banco.ConsultaDataSet(Sql, "Carteiras")
        ddlCarteiras.DataBind()

        ddlCarteiras.Items.Insert(0, "")
        ddlCarteiras.SelectedIndex = 0

    End Sub

    Sub CargaBancoRecebedor()
        DdlBancoRecebedor.Items.Clear()
        DdlContaRecebedora.Items.Clear()

        If DdlEmpresaPagadora.SelectedValue <> "" Then
            Cliente = DdlEmpresaPagadora.SelectedValue
            campo = Cliente.Split("-")

            Sql = " SELECT DISTINCT BxC.Banco_Id,  B.Descricao" & vbCrLf & _
                  "   FROM BancosXContas BxC " & vbCrLf & _
                  "  INNER JOIN Bancos B " & vbCrLf & _
                  "     ON BxC.Banco_Id = B.Banco_Id" & vbCrLf & _
                  "  WHERE BxC.Empresa_Id  = '" & campo(0) & "'" & vbCrLf & _
                  "    AND BxC.EndEmpresa_Id  = " & campo(1) & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Bancos").Tables(0).Rows
                Descricao = Format(Dr("Banco_Id"), "0000") & "- " & Dr("Descricao")
                DdlBancoRecebedor.Items.Add(New ListItem(Descricao, Dr("Banco_Id")))
            Next
        End If

        DdlBancoRecebedor.Items.Insert(0, "")
        DdlBancoRecebedor.SelectedIndex = 0
    End Sub

    Function ValidaCampos() As Boolean
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        'Empresa Recebedora para utilizar na validação de datas não programáveis
        Dim Emp() As String = Nothing
        If DdlEmpresaPagadora.SelectedIndex > 0 Then
            Emp = DdlEmpresaPagadora.SelectedValue.Split("-")
        Else
            Emp = DdlEmpresaCliente.SelectedValue.Split("-")
        End If

        Dim Mensagem As String = String.Empty

        If Not ValidaData(txtMovimento.Text, "Movimento", Emp(0), Emp(1)) Then
            Return False
        ElseIf DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = 0 And lblAgrupar.Text <> "BT" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório ...")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresaCliente.Text) And lblAgrupar.Text <> "BT" Then
            MsgBox(Me.Page, "Empresa do Cliente é obrigatório ...")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCliente.Text) And lblAgrupar.Text <> "BT" Then
            MsgBox(Me.Page, "Cliente é obrigatório ...")
            Return False
        ElseIf ddlIndexador.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Indexador é obrigatório.")
            Return False
        ElseIf txtPedido.Text.Length > 0 AndAlso ddlMoeda.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Titulo com Pedido a Moeda deve ser selecionada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlProvisoes.Text) Then
            MsgBox(Me.Page, "Previsao é obrigatório ...")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 3 And Not Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
            MsgBox(Me.Page, "Usuário sem autorização para lançamento de Provisão")
            Return False
        ElseIf txtHistorico.Text = "" And lblAgrupar.Text <> "BT" Then
            MsgBox(Me.Page, "Histórico é obrigatório ...")
            Return False
        ElseIf ddlCarteiras.SelectedIndex = 0 AndAlso Not lblAgrupar.Text = "AP" Then
            MsgBox(Me.Page, "Finalidade Financeira é obrigatório ...")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 1 AndAlso ddlCarteiras.SelectedIndex > 0 AndAlso DdlTributos.Items.Count > 1 AndAlso DdlTributos.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Encargo é obrigatório ...")
            Return False
            'ElseIf Not ValidaData(txtProrrogacao.Text, "Vencimento", Session("ssEmpresa"), Session("ssEndEmpresa")) Then
            '    Return False
        ElseIf Not ValidaData(txtProrrogacao.Text, "Vencimento", Emp(0), Emp(1)) Then
            MsgBox(Me.Page, Mensagem)
            Return False
        End If

        Valor = CDbl(txtValorDoDocumento.Text) + CDbl(txtValorEmMoeda.Text)

        If Valor = 0 Then
            txtValorDoDocumento.Text = ""
            txtDescontos.Text = ""
            txtDeducoes.Text = ""
            txtJuros.Text = ""
            txtAcrescimos.Text = ""
            txtValorCobrado.Text = ""
            MsgBox(Me.Page, "Valor Do Documento em R$ ou U$ é obrigatório ...")
            Return False
        End If

        If Not Funcoes.VerificaPermissao("AJUSTEFINANCEIRO", "GRAVAR") Then
            If (String.IsNullOrWhiteSpace(txtValorDoDocumento.Text) AndAlso String.IsNullOrWhiteSpace(txtValorEmMoeda.Text)) Then
                MsgBox(Me.Page, "Valor Do Documento em R$ ou U$ é obrigatório e maior que zero!")
                Return False
            ElseIf (ddlMoeda.SelectedValue = 1 AndAlso (String.IsNullOrWhiteSpace(txtValorDoDocumento.Text) OrElse CDbl(txtValorDoDocumento.Text) <= 0)) Then
                MsgBox(Me.Page, "Valor Do Documento em R$ é obrigatório.")
                Return False
            ElseIf (ddlMoeda.SelectedValue = 3 AndAlso (String.IsNullOrWhiteSpace(txtValorEmMoeda.Text) OrElse CDbl(txtValorEmMoeda.Text) <= 0)) Then
                MsgBox(Me.Page, "Valor Do Documento em U$ é obrigatório.")
                Return False
            ElseIf DdlProvisoes.SelectedValue = 1 AndAlso txtValorDoDocumento.Text = 0 Then
                MsgBox(Me.Page, "Valor Em Reais é Obrigatório ...")
                Return False
            ElseIf DdlProvisoes.SelectedValue = 1 And txtValorCobrado.Text = "" And lblAgrupar.Text <> "BT" Then
                MsgBox(Me.Page, "Valor Pago é obrigatório ...")
                Return False
            End If
        End If

        If txtLiberarTitulo.Value = "N" And txtPedido.Text <> "" And txtPedido.Text > 0 Then
            If ddlMoeda.SelectedValue = 1 Then
                If txtValorDocumento.Value <> 0 Then
                    If CDbl(txtValorDoDocumento.Text) > txtValorDocumento.Value + 1 Then
                        MsgBox(Me.Page, "Valor do documento não pode ser maior do que o valor programado pelo Pedido ...")
                        Return False
                    End If
                End If
            ElseIf ddlMoeda.SelectedValue = 3 Then
                If txtValorMoeda.Value <> 0 Then
                    If CDbl(txtValorEmMoeda.Text) > txtValorMoeda.Value + 1 Then
                        MsgBox(Me.Page, "Valor do documento não pode ser maior do que o valor programado pelo Pedido ...")
                        Return False
                    End If
                End If
            End If
        End If

        If DdlProvisoes.SelectedValue = 1 Then 'em caso de baixa.
            If String.IsNullOrWhiteSpace(DdlEmpresaPagadora.Text) Then
                MsgBox(Me.Page, "Empresa Recebedora é obrigatório ...")
                Return False
            ElseIf Not String.IsNullOrWhiteSpace(DdlEmpresaPagadora.SelectedValue) AndAlso String.IsNullOrWhiteSpace(DdlTiposDePagamentos.Text) Then
                MsgBox(Me.Page, "Tipo de Recebimento é obrigatório ...")
                Return False
            ElseIf DdlTiposDePagamentos.Text <> "" AndAlso DdlBancoRecebedor.Text = "" Then
                MsgBox(Me.Page, "Banco é obrigatório ...")
                Return False
            ElseIf DdlTiposDePagamentos.Text <> "" AndAlso DdlContaRecebedora.Text = "" Then
                MsgBox(Me.Page, "Conta Bancária é obrigatório ...")
                Return False
            End If
        End If


        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        Dim Empresa As New [Lib].Negocio.ClienteXEmpresa(campo(0), campo(1))

        If Not String.IsNullOrWhiteSpace(Empresa.CodigoContaFornecedorFrete) _
            AndAlso Not lblAgrupar.Text = "AP" AndAlso String.IsNullOrWhiteSpace(txtRegistro.Text) _
            AndAlso Empresa.CodigoContaFornecedorFrete = objCarteira.CodigoContaCliente Then
            MsgBox(Me.Page, "Titulo não pode ser incluído na Conta de Fornecedor de Frete.")
            Return False
        End If

        '******************************************************************************************************
        '***************************************** Adiantamento ***********************************************
        '******************************************************************************************************
        If objCarteira.isAdiantamento AndAlso objCarteira.BaixaAdiantamento AndAlso (Not IsNumeric(txtNumeroAdto.Text) OrElse CInt(txtNumeroAdto.Text) <= 0) Then
            MsgBox(Me.Page, "Selecione um adiantamento para efetuar a baixa")
            Return False
        End If

        If Not objCarteira.isAdiantamento AndAlso IsNumeric(txtNumeroAdto.Text) AndAlso CInt(txtNumeroAdto.Text) > 0 Then
            MsgBox(Me.Page, "Numero de adiantamento nao pode ser informado com carteira que nao sao de adiantamento ou baixa. reinicie o lançamento.")
            Return False
        End If

        If objCarteira.isAdiantamento AndAlso Not objCarteira.BaixaAdiantamento AndAlso String.IsNullOrWhiteSpace(txtVencimentoAdto.Text) Then
            MsgBox(Me.Page, "Vencimento para o Adiantamento não foi informado, Verifique.")
            Return False
        End If

        If objCarteira.isAdiantamento And objCarteira.BaixaAdiantamento Then
            If ddlMoeda.SelectedValue = 1 AndAlso CDec(HDSaldoAdiantamento.Value) < CDec(txtValorDoDocumento.Text) Then
                MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                Return False
            ElseIf Not ddlMoeda.SelectedValue = 1 AndAlso CDec(HDSaldoAdiantamento.Value) < CDec(txtValorEmMoeda.Text) Then
                MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                Return False
            End If
        End If

        '******************************************************************************************************
        '******************************************************************************************************

        If DdlTributos.SelectedIndex > 0 Then 'Consulta Contas Contabeis dos Tributos
            Dim Encargo As New Encargo(DdlTributos.SelectedValue)
            If String.IsNullOrEmpty(Encargo.ContaDebito) And String.IsNullOrEmpty(Encargo.ContaCredito) Then
                MsgBox(Me.Page, "Encargo Sem contas Credito e Debito Cadastradas, Verifique o encargo.")
                Return False
            ElseIf Not String.IsNullOrWhiteSpace(Empresa.CodigoContaFornecedorFrete) _
                AndAlso String.IsNullOrWhiteSpace(txtRegistro.Text) _
                AndAlso (Empresa.CodigoContaFornecedorFrete = Encargo.ContaDebito OrElse Empresa.CodigoContaFornecedorFrete = Encargo.ContaCredito) Then
                MsgBox(Me.Page, "Titulo não pode ser incluído na Conta de Fornecedor de Frete.")
                Return False
            End If
        End If

        If Not lblAgrupar.Text = "AP" AndAlso DdlTributos.SelectedIndex = 0 AndAlso Not objCarteira.CodigoContaCliente Is Nothing AndAlso objCarteira.CodigoContaCliente.Trim.Length = 0 Then
            ddlCarteiras.SelectedIndex = 0
            MsgBox(Me.Page, "Carteira Financeira Sem Conta Contábil, Verifique.")
            Return False
        ElseIf txtRegistro.Text.Length > 0 And txtPedido.Text <> "0" Then
            Dim crtl As String() = Session("ControleCR" & HID.Value).ToString.Split(";")

            If crtl(0) <> txtRegistro.Text Then
                MsgBox(Me.Page, "Numero de registro de verificação invalido.")
                Return False
            ElseIf New [Lib].Negocio.Moeda(crtl(3)).Classificacao = eTiposMoeda.Oficial Then
                If CDec(txtValorDoDocumento.Text) > CDec(crtl(1)) Then
                    MsgBox(Me.Page, "Valor do Documento R$ não pode ser maior que o original: " & CDec(crtl(1)).ToString("N2"))
                    Return False
                End If
            Else
                If CDec(txtValorEmMoeda.Text) > CDec(crtl(2)) Then
                    MsgBox(Me.Page, "Valor do Documento U$ não pode ser maior que o original: " & CDec(crtl(2)).ToString("N2"))
                    Return False
                End If
            End If
        End If

        If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 AndAlso objCarteira.isAdiantamento Then
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
            If objCliente.DesdobrarFornecedor = True Then
                MsgBox(Me.Page, "Fornecedor do Desdobramento do Título não foi informado")
                Return False
            End If
        End If

        If txtRegistro.Text.Length = 0 And txtPedido.Text <> "0" Then
            Dim objpedido As New [Lib].Negocio.Pedido(campo(0), campo(1), txtPedido.Text)
            If Trim(txtValorEmMoeda.Text) = "" Then txtValorEmMoeda.Text = 0 'carrega dolar caso nao tenha sido preenchido para evitar erros nas funcoes abaixo.
        End If

        'Verificar se as informações como Unidade, Empresa e Cliente não tenham sido alteradas no "meio tempo" entre a consulta e baixa do título
        If txtPedido.Text <> "0" Then
            'Dim sql As String
            Sql = "SELECT Empresa_Id, EndEmpresa_Id, UnidadeDeNegocio, Cliente, EndCliente " & vbCrLf & _
                  "  FROM Pedidos " & vbCrLf & _
                  " WHERE pedido_id = " & txtPedido.Text & vbCrLf & _
                  "   AND UnidadeDeNegocio = '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'" & vbCrLf & _
                  "   AND Empresa_Id= '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & vbCrLf & _
                  "   AND Cliente =  '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                  "   AND EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pedido")

            If ds.Tables("Pedido").Rows.Count <= 0 Then
                MsgBox(Me.Page, "Recarregue o título, pois houveram alterações em sua Nota ou Pedido.")
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub ValidaValores(ByVal dolarizar As Boolean)
        Dim Zero As Decimal = 0

        txtValorDoDocumento.Text = Funcoes.AtribuirirValorFormatado(txtValorDoDocumento.Text)
        txtDescontos.Text = Funcoes.AtribuirirValorFormatado(txtDescontos.Text)
        txtJuros.Text = Funcoes.AtribuirirValorFormatado(txtJuros.Text)
        txtAcrescimos.Text = Funcoes.AtribuirirValorFormatado(txtAcrescimos.Text)
        txtDeducoes.Text = Funcoes.AtribuirirValorFormatado(txtDeducoes.Text)
        txtValorCobrado.Text = Funcoes.AtribuirirValorFormatado(txtValorCobrado.Text)
        txtValorEmMoeda.Text = Funcoes.AtribuirirValorFormatado(txtValorEmMoeda.Text)
        txtDescontosMoeda.Text = Funcoes.AtribuirirValorFormatado(txtDescontosMoeda.Text)
        txtJurosMoeda.Text = Funcoes.AtribuirirValorFormatado(txtJurosMoeda.Text)
        txtAcrescimosMoeda.Text = Funcoes.AtribuirirValorFormatado(txtAcrescimosMoeda.Text)
        txtDeducoesMoeda.Text = Funcoes.AtribuirirValorFormatado(txtDeducoesMoeda.Text)
        txtValorCobradoMoeda.Text = Funcoes.AtribuirirValorFormatado(txtValorCobradoMoeda.Text)

        If dolarizar Then
            If ddlMoeda.SelectedValue = 1 AndAlso CDec(txtValorEmMoeda.Text) = 0 Then
                Dim vlr As Decimal = DolarizaBaixa(txtValorDoDocumento.Text, 2)
                txtValorEmMoeda.Text = vlr.ToString("N2")
            ElseIf ddlMoeda.SelectedValue = 3 AndAlso (DdlProvisoes.Text = 2 Or DdlProvisoes.Text = 3) Then
                If txtPedido.Text > 0 Then
                    Sql = "Select DataPedido, isnull(IndiceFixado,0) as IndiceFixado From Pedidos where Pedido_id = " & txtPedido.Text
                    Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql, "Pedidos")

                    If dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                        txtValorCobrado.Text = Math.Round(CDbl(txtValorEmMoeda.Text) * dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                    Else
                        If Not ddlIndexador.SelectedValue = 99 Then txtValorCobrado.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtValorEmMoeda.Text), ddlIndexador.SelectedValue, dsPedido.Tables(0).Rows(0).Item("DataPedido")), 2)
                    End If

                    If CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text) < 0 Then
                        txtAcrescimos.Text = ((CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text)) * -1).ToString("N2")
                    Else
                        txtDeducoes.Text = (CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text)).ToString("N2")
                    End If
                    dolarizar = False
                End If
            End If

            If ddlMoeda.SelectedValue = 1 Then
                If ddlIndexador.SelectedValue = 99 Then
                    If CDec(txtDescontos.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtDescontosMoeda.Text = (CDec(txtDescontos.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text))).ToString("N2")
                    Else
                        txtDescontosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtDeducoes.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtDeducoesMoeda.Text = Math.Round(CDec(txtDeducoes.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)), 2, MidpointRounding.AwayFromZero)
                    Else
                        txtDeducoesMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtJuros.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtJurosMoeda.Text = (CDec(txtJuros.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text))).ToString("N2")
                    Else
                        txtJurosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtAcrescimos.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtAcrescimosMoeda.Text = (CDec(txtAcrescimos.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text))).ToString("N2")
                    Else
                        txtAcrescimosMoeda.Text = Zero.ToString("N2")
                    End If
                Else
                    If CDec(txtDescontos.Text) > 0 Then
                        txtDescontosMoeda.Text = DolarizaBaixa(txtDescontos.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtDescontosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtDeducoes.Text) > 0 Then
                        txtDeducoesMoeda.Text = DolarizaBaixa(txtDeducoes.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtDeducoesMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtJuros.Text) > 0 Then
                        txtJurosMoeda.Text = DolarizaBaixa(txtJuros.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtJurosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtAcrescimos.Text) > 0 Then
                        txtAcrescimosMoeda.Text = DolarizaBaixa(txtAcrescimos.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtAcrescimosMoeda.Text = Zero.ToString("N2")
                    End If
                End If
            ElseIf ddlMoeda.SelectedValue = 3 Then
                If CDec(txtDescontosMoeda.Text) > 0 Then
                    txtDescontos.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtDescontosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtDescontos.Text = Zero.ToString("N2")
                End If
                If CDec(txtDeducoesMoeda.Text) > 0 Then
                    txtDeducoes.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtDeducoesMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtDeducoes.Text = Zero.ToString("N2")
                End If
                If CDec(txtJurosMoeda.Text) > 0 Then
                    txtJuros.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtJurosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtJuros.Text = Zero.ToString("N2")
                End If
                If CDec(txtAcrescimosMoeda.Text) > 0 Then
                    txtAcrescimos.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtAcrescimosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtAcrescimos.Text = Zero.ToString("N2")
                End If
            End If
        End If

        txtValorCobrado.Text = Math.Round(CDec(txtValorDoDocumento.Text) + CDec(txtJuros.Text) + CDec(txtAcrescimos.Text) - CDec(txtDescontos.Text) - CDec(txtDeducoes.Text), 2, MidpointRounding.AwayFromZero)
        txtValorCobradoMoeda.Text = Math.Round(CDec(txtValorEmMoeda.Text) + CDec(txtJurosMoeda.Text) + CDec(txtAcrescimosMoeda.Text) - CDec(txtDescontosMoeda.Text) - CDec(txtDeducoesMoeda.Text), 2, MidpointRounding.AwayFromZero)

        If lnkNovo.Parent.Visible = False Then lnkNovo.Parent.Visible = True
    End Sub

    Function LanctosContabeis()

        Sql = "INSERT INTO Razao " & vbCrLf & _
              "       (Empresa_Id, " & vbCrLf & _
              "       EndEmpresa_Id, " & vbCrLf & _
              "       Conta_Id, " & vbCrLf & _
              "       Cliente_Id, " & vbCrLf & _
              "       EndCliente_Id, " & vbCrLf & _
              "       Movimento_Id, " & vbCrLf & _
              "       Lote_Id, " & vbCrLf & _
              "       Sequencia_Id, " & vbCrLf & _
              "       Titulo, " & vbCrLf & _
              "       UnidadeDeNegocio, " & vbCrLf & _
              "       Indexador, " & vbCrLf & _
              "       DataMoeda, " & vbCrLf & _
              "       DebitoOficial, " & vbCrLf & _
              "       CreditoOficial, " & vbCrLf & _
              "       DebitoMoeda, " & vbCrLf & _
              "       CreditoMoeda, " & vbCrLf & _
              "       Historico, " & vbCrLf & _
              "       PrevistoRealizado," & vbCrLf & _
              "       Processo," & vbCrLf & _
              "       UsuarioInclusao," & vbCrLf & _
              "       UsuarioInclusaoData)" & vbCrLf & _
              "       VALUES (" & vbCrLf & _
              "               '" & Raz_Empresa & "'," & vbCrLf & _
              "                " & Raz_EndEmpresa & "," & vbCrLf & _
              "               '" & Raz_Conta & "'" & vbCrLf

        If Len(Raz_Conta) = 7 Then
            Sql &= ", '" & Raz_Cliente & "'"        'Cliente
            Sql &= ", " & Raz_EndCliente            'Endereco do Cliente
        Else
            Sql &= ", ''"                           'Cliente
            Sql &= ", 0"                            'Endereco do Cliente
        End If

        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'"     'Data de Movimento
        Sql &= ", 0070"
        Sql &= ", " & Registro                      'Sequencia no Razao = Registro do Titulo
        Sql &= ", " & Registro                      'Numero do Titulo
        Sql &= ", '" & Raz_UnidadeDeNegocio & "'"   'Unidade de Negócio
        Sql &= ", " & ddlIndexador.SelectedValue    'Indexador
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'"     'Data da Moeda

        'Valor Oficial
        If Raz_DebitoCredito = "D" Then
            Sql &= ", " & Replace(Raz_ValorOficial, ",", ".")  'Valor Débito Oficial
            Sql &= ", 0.0"                                     'Valor Crédito Oficial
        Else
            Sql &= ", 0.0"                                     'Valor Debito Oficial
            Sql &= ", " & Replace(Raz_ValorOficial, ",", ".")  'Valor Crédito Oficial
        End If

        If IsVariacao(DdlEmpresaCliente.SelectedValue, Raz_Conta) Then
            Sql &= ", 0.0"      'Valor Débito Moeda
            Sql &= ", 0.0"      'Valor Crédito Moeda
        Else
            If Raz_DebitoCredito = "D" Then
                Sql &= ", " & Replace(Raz_ValorMoeda, ",", ".")  'Valor Débito Moeda
                Sql &= ", 0.0"                                   'Valor Crédito Moeda
            Else
                Sql &= ", 0.0"                                   'Valor Debito Moeda
                Sql &= ", " & Replace(Raz_ValorMoeda, ",", ".")  'Valor Crédito Moeda
            End If
        End If

        If Raz_Historico.Length > 200 Then
            Raz_Historico = Raz_Historico.Substring(0, 200)
        End If

        Sql &= ", '" & Raz_Historico & "'"          'Histórico
        Sql &= ", 'P'"                              'Previsto/Realizado
        Sql &= ", 'CONTASARECEBER'"                 'Processo
        Sql &= ", '" & Session("ssNomeUsuario") & "'"  'Usuario que Baixou
        Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')"           'Data da Baixa

        SqlArray.Add(Sql)

        Return True

    End Function

    Private Function IsVariacao(ByVal Empresa As String, ByVal Conta As String) As Boolean
        Dim sql As String = "SELECT CASE" & vbCrLf & _
                            "            WHEN '" & Conta & "' in (empDed.contaVariacaoMonetariaAtiva, empDed.contaVariacaoMonetariaPassiva) " & vbCrLf & _
                            "                THEN 1" & vbCrLf & _
                            "                ELSE 0" & vbCrLf & _
                            "       END Variacao" & vbCrLf & _
                            "  FROM ClientesxEmpresas empDed" & vbCrLf & _
                            " WHERE empDed.empresa_id    = '" & Empresa.Split("-")(0) & "'" & vbCrLf & _
                            "   AND empDed.endempresa_id =  " & Empresa.Split("-")(1)

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Deducoes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                If row("Variacao") = 1 Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Sub GravaTitulo()
        Dim Cliente As String
        Dim Campo() As String
        Dim Valor As String

        ValidaValores(False)

        If Not ValidaCampos() Then Exit Sub
        '''''''''''''''''''''''''''''''
        'Gera sequencia de titulos'''
        ''''''''''''''''''''''''''''''
        If String.IsNullOrWhiteSpace(txtRegistro.Text) Then
            Sql = "exec sp_Numerador '" & Session("ssNomeServidor").ToUpper() & "',0,1"
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
                Registro = Dr("Sequencia")
                txtRegistro.Text = Registro
            Next
        Else
            Registro = CInt(txtRegistro.Text)

            Sql = "DELETE TitulosXDesdobrarFornecedor " & vbCrLf & _
                  "WHERE Registro_Id = " & Registro
            SqlArray.Add(Sql)

            '---Adiantamentos X Baixas----------------
            Sql = "DELETE FROM AdiantamentosXBaixas" & vbCrLf & _
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            '---Adiantamentos -----------------------
            Sql = "DELETE FROM Adiantamentos" & vbCrLf & _
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE FROM razao" & vbCrLf & _
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE FROM ContasAReceber" & vbCrLf & _
                  " WHERE Registro_Id = " & Registro
            SqlArray.Add(Sql)
        End If

        Data = txtMovimento.Text

        'Dolariza()

        Sql = "INSERT INTO ContasAReceber " & vbCrLf & _
              "       (Registro_Id" & vbCrLf & _
              "       ,Sequencia_Id" & vbCrLf & _
              "       ,Provisao" & vbCrLf & _
              "       ,Carteira" & vbCrLf & _
              "       ,Tributo" & vbCrLf & _
              "       ,Indexador" & vbCrLf & _
              "       ,Moeda" & vbCrLf & _
              "       ,TipoPagto" & vbCrLf & _
              "       ,Situacao" & vbCrLf & _
              "       ,Lote" & vbCrLf & _
              "       ,Movimento" & vbCrLf & _
              "       ,Vencimento" & vbCrLf & _
              "       ,Prorrogacao" & vbCrLf & _
              "       ,DataMoeda" & vbCrLf & _
              "       ,Baixa" & vbCrLf & _
              "       ,UnidadeDeNegocio" & vbCrLf & _
              "       ,Empresa" & vbCrLf & _
              "       ,EndEmpresa" & vbCrLf & _
              "       ,Cliente" & vbCrLf & _
              "       ,EndCliente" & vbCrLf & _
              "       ,BancoCliente" & vbCrLf & _
              "       ,AgenciaCliente" & vbCrLf & _
              "       ,DigitoAgenciaCliente" & vbCrLf & _
              "       ,ContaCliente" & vbCrLf & _
              "       ,DigitoContaCliente" & vbCrLf & _
              "       ,ContaContabilCliente" & vbCrLf & _
              "       ,EmpresaPagadora" & vbCrLf & _
              "       ,EndEmpresaPagadora" & vbCrLf & _
              "       ,BancoPagador" & vbCrLf & _
              "       ,AgenciaPagadora" & vbCrLf & _
              "       ,DigitoAgenciaPagadora" & vbCrLf & _
              "       ,ContaPagadora" & vbCrLf & _
              "       ,DigitoContaPagadora" & vbCrLf & _
              "       ,ContaContabilPagadora" & vbCrLf & _
              "       ,Cheque" & vbCrLf & _
              "       ,Slips" & vbCrLf & _
              "       ,Recibo" & vbCrLf & _
              "       ,Aviso" & vbCrLf & _
              "       ,ReciboDeposito" & vbCrLf & _
              "       ,EmpresaPedido" & vbCrLf & _
              "       ,EndEmpresaPedido" & vbCrLf & _
              "       ,Pedido" & vbCrLf & _
              "       ,PedidoFixacao" & vbCrLf & _
              "       ,ValorDoDocumento" & vbCrLf & _
              "       ,Descontos" & vbCrLf & _
              "       ,Deducoes" & vbCrLf & _
              "       ,Juros" & vbCrLf & _
              "       ,Acrescimos" & vbCrLf & _
              "       ,ValorLiquido" & vbCrLf & _
              "       ,MoedaValorDoDocumento" & vbCrLf & _
              "       ,MoedaDescontos" & vbCrLf & _
              "       ,MoedaDeducoes" & vbCrLf & _
              "       ,MoedaJuros" & vbCrLf & _
              "       ,MoedaAcrescimos" & vbCrLf & _
              "       ,MoedaValorLiquido" & vbCrLf & _
              "       ,Historico" & vbCrLf & _
              "       ,CodigoDeBarras" & vbCrLf & _
              "       ,CodigoDigitado" & vbCrLf & _
              "       ,Destinatario" & vbCrLf & _
              "       ,EndDestinatario" & vbCrLf & _
              "       ,solicitacao" & vbCrLf & _
              "       ,UsuarioInclusao" & vbCrLf & _
              "       ,UsuarioInclusaoData" & vbCrLf & _
              "       ,UsuarioAlteracao" & vbCrLf & _
              "       ,UsuarioAlteracaoData" & vbCrLf & _
              "       ,UsuarioBaixa" & vbCrLf & _
              "       ,UsuarioBaixaData" & vbCrLf & _
              "       ,Grupado" & vbCrLf & _
              "       ,Observacoes" & vbCrLf & _
              "       ,SituacaoBancaria" & vbCrLf & _
              "       ,UsuarioLiberacaoBloqueio" & vbCrLf & _
              "       ,UsuarioLiberacaoBloqueioDate" & vbCrLf & _
              "       ,UsuarioLiberacaoPedido" & vbCrLf & _
              "       ,UsuarioLiberacaoPedidoDate" & vbCrLf & _
              "       ,CarteiraDoTitulo" & vbCrLf & _
              "       ,ContratoDeFinanciamento" & vbCrLf & _
              "       ,ContratoBancario)" & vbCrLf & _
              " VALUES( " & vbCrLf

        Sql &= Registro                  'Registro
        Sql &= ", 0"                      'Sequencia
        Sql &= ", " & DdlProvisoes.SelectedValue
        Sql &= ", '" & ddlCarteiras.SelectedValue & "'"

        If DdlTributos.Text <> "" Then
            Sql &= ", '" & DdlTributos.SelectedValue & "'"
        Else
            Sql &= ", ''"
        End If

        Sql &= ", " & ddlIndexador.SelectedValue             'Indexador
        Sql &= ", " & ddlMoeda.SelectedValue                 'Moeda

        If DdlTiposDePagamentos.Text <> "" Then
            Sql &= ", " & DdlTiposDePagamentos.SelectedValue
        Else
            Sql &= ", 0"
        End If

        Sql &= ", 1"                        'Situacao
        Sql &= ", 70"                       'Lote
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" 'movimento
        If lblVencOriginal.Text.Length = 0 Then
            Sql &= ", '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" 'Vencimento
        Else
            Sql &= ", '" & CDate(lblVencOriginal.Text).ToString("yyyy/MM/dd") & "'" 'vencimento
        End If
        Sql &= ", '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" 'prorrogacao
        Sql &= ", '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" 'data moeda
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" 'data baixa

        Sql &= ", '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'"

        Cliente = DdlEmpresaCliente.SelectedValue
        Campo = Cliente.Split("-")

        If Campo(0) <> "" Then
            Sql &= ", '" & Campo(0) & "'"                           'EmpresaCliente
            Sql &= ", " & Campo(1)                                  'Endereco Empresa Cliente
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If

        Cliente = txtCodigoCliente.Value
        Campo = Cliente.Split("-")

        If Campo(0) <> "" Then
            Sql &= ", '" & Campo(0) & "'"                           'Cliente
            Sql &= ", " & Campo(1)                                  'Endereco Cliente
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If

        '-Dados Bancarios Cliente----------

        Sql &= ", 0"                        'Banco Cliente
        Sql &= ", ''"                       'Agencia do Destinatario
        Sql &= ", ''"                       'Digito da Agencia do Destinatário
        Sql &= ", ''"                       'Conta Corrente do Destinatário
        Sql &= ", ''"                       'Digito da Conta Corrente do Destinatário

        If DdlTributos.Text <> "" Then
            Sql &= ", '" & Carteira(DdlTributos.SelectedValue) & "'"  'Grupo de Conta Corrente do Fornecedor
        Else
            Sql &= ", '" & Carteira(ddlCarteiras.SelectedValue) & "'"  'Grupo de Conta Corrente do Fornecedor
        End If

        'Empresa Pagadora------------------

        If DdlTiposDePagamentos.Text <> "" AndAlso DdlTiposDePagamentos.Text <> "0" AndAlso DdlProvisoes.SelectedValue = 1 Then
            Cliente = DdlEmpresaPagadora.SelectedValue
            Campo = Cliente.Split("-")
            Sql &= ", '" & Campo(0) & "'"                           'Empresa Recebedora
            Sql &= ", " & Campo(1)                                  'Endereco Empresa Recebedora

            Sql &= ", '" & DdlBancoRecebedor.SelectedValue & "'"    'Banco Cliente
            Cliente = DdlContaRecebedora.SelectedValue
            Campo = Cliente.Split("-")
            Sql &= ", '" & Campo(0) & "'"                           'Agencia Empresa
            Sql &= ", '" & Campo(1) & "'"                           'Digito Agencia 
            Sql &= ", '" & Campo(2) & "'"                           'Conta Empresa
            Sql &= ", '" & Campo(3) & "'"                           'Digito Conta Empresa
            Sql &= ", '" & Campo(4) & "'"                           'Conta Contabil
        Else
            Sql &= ", ''"                                           'Empresa Recebedora
            Sql &= ", 0"                                            'Endereco Empresa Recebedora
            Sql &= ", 0"                                            'Banco Recebedor
            Sql &= ", ''"                                           'Agencia Recebedora
            Sql &= ", ''"                                           'Digito Agencia Recebedora
            Sql &= ", ''"                                           'Conta Corrente Recebedora
            Sql &= ", ''"                                           'Digito Conta Recebedora
            Sql &= ", ''"                                           'Digito Conta Recebedora
        End If

        '-----------------------------------------------

        Sql &= ", 'N'"                                              'Emite Cheque
        Sql &= ", 'N'"                                              'Emite Slips
        Sql &= ", 'N'"                                              'Emite Recibo
        Sql &= ", 'N'"                                              'Emite Aviso
        Sql &= ", 'N'"                                              'Emite Recibo De Deposito

        If String.IsNullOrWhiteSpace(txtPedido.Text) OrElse CInt(txtPedido.Text) = "0" Then
            Sql &= ", NULL"
            Sql &= ", NULL"
            Sql &= ", NULL"
            Sql &= ", NULL"
        Else
            Cliente = DdlEmpresaCliente.SelectedValue
            Campo = Cliente.Split("-")
            Sql &= ", '" & Campo(0) & "'"
            Sql &= ", " & Campo(1)
            If IsNumeric(txtPedido.Text) AndAlso CInt(txtPedido.Text) = 0 Then
                Sql &= ", NULL"
            Else
                Sql &= ", " & txtPedido.Text
            End If

            If txtPedidoFixacao.Value.ToString.Length = 0 OrElse txtPedidoFixacao.Value = 0 Then  'PedidoFixacao
                Sql &= ", NULL"
            Else
                Sql &= ", " & txtPedidoFixacao.Value
            End If
        End If

        '---Valores em Reais ----------------------------
        'Valor do Documento
        Valor = Replace(txtValorDoDocumento.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Descontos
        Valor = Replace(txtDescontos.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Deducoes
        Valor = Replace(txtDeducoes.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Juros
        Valor = Replace(txtJuros.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Acrescimos
        Valor = Replace(txtAcrescimos.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Liquido
        Valor = Replace(txtValorCobrado.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")

        '---Valores em Dolar------------------
        'Valor do Documento
        Valor = Replace(txtValorEmMoeda.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Descontos
        Valor = Replace(txtDescontosMoeda.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Deducoes
        Valor = Replace(txtDeducoesMoeda.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Juros
        Valor = Replace(txtJurosMoeda.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Acrescimos
        Valor = Replace(txtAcrescimosMoeda.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")
        'Liquido
        Valor = Replace(txtValorCobradoMoeda.Text, ".", "")
        Sql &= ", " & Replace(Valor, ",", ".")

        Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text) & "'"                  'Historico
        Sql &= ", ''"                                           'Codigo De Barras
        Sql &= ", 'N'"                                          'Digitado

        Cliente = txtCodigoCliente.Value
        Campo = Cliente.Split("-")

        If Campo(0) <> "" Then
            Sql &= ", '" & Campo(0) & "'"                          'Cliente 
            Sql &= ", " & Campo(1)                                 'Endereco Cliente 
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If

        Sql &= ", 0"                                                                'Solicitacao

        If lblUsuarioIncl.Text.Length = 0 Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"           'Usuario Que Esta Incluindo
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                     'Data da desta Inclusao
            Sql &= ", ''"                                                               'UsuarioAlteracao
            Sql &= ", ''"                                                               'UsuarioData
        Else
            Dim Usu As Array = lblUsuarioIncl.Text.Trim.Split("-")
            Sql &= ", '" & Usu(0) & "'"                                                  'Usuario que Incluiu
            Sql &= ", '" & CDate(Usu(1)).ToString("yyyy/MM/dd") & "'"                     'Data De Quando Ocorreu a Inclusao
            Sql &= ", '" & Session("ssNomeUsuario") & "'"            'Usuario Que Esta alterando
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                      'Data Desta Alteracao
        End If

        If DdlProvisoes.Text <> "" Then
            If DdlProvisoes.SelectedValue = 1 Then
                Sql &= ", '" & Session("ssNomeUsuario") & "'"   'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"             'Data da Baixa
            Else
                Sql &= ", ''"
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"             'Data da Baixa
            End If
        Else
            Sql &= ", ''"
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                 'Data da Baixa
        End If

        If lblAgrupar.Text = "AP" Then
            Sql &= ", 'M'"                                                          'Registro Grupado
            Sql &= ", '" & Session("ssObservacoes" & HID.Value) & "'"      'Observaçoes de Agrupamento
        Else
            Sql &= ", 'N'"                                                          'Registro Grupado
            Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text) & " '"                               'Observacoes
        End If

        Sql &= ", 0"                      'Situacao Bancaria


        If txtLiberarTitulo.Value = "S" Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"        'Usuario que Liberou Titulo
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                  'Data da Liberação do Titulo
        Else
            If txtUsuarioLiberarTitulo.Value = "" Then
                Sql &= ", ''"                                                        'Usuario que Liberou Titulo
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"              'Data da Liberação do Titulo
            Else
                Sql &= ", '" & txtUsuarioLiberarTitulo.Value & "'"                                  'Usuario que Liberou Titulo
                Sql &= ", '" & CDate(txtUsuarioLiberarTituloData.Value).ToString("yyyy/MM/dd") & "'" 'Data da Liberação do Titulo
            End If
        End If

        If txtLiberarPedido.Value = "S" Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"        'Usuario que Liberou Pedido
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                  'Data da Liberação do Pedido
        Else
            If txtUsuarioLiberarPedido.Value = "" Then
                Sql &= ", ''"                                                        'Usuario que Liberou Pedido
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"              'Data da Liberação do Pedido
            Else
                Sql &= ", '" & txtUsuarioLiberarPedido.Value & "'"                                  'Usuario que Liberou Pedido
                Sql &= ", '" & CDate(txtUsuarioLiberarPedidoData.Value).ToString("yyyy/MM/dd") & "'" 'Data da Liberação do Pedido
            End If
        End If

        Sql &= ", " & ddlCarteiraDoTitulo.SelectedValue & ""                         'Carteira do Titulo 
        Sql &= ", '" & txtContratoFinanceiro.Text & "'"                              'Contrato De Financiamento 
        Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(txtContratoBanco.Text) & "')"  'Contrato Bancário 

        SqlArray.Add(Sql)

        '--Grava Novo Titulo de Pagamentos Parciais------------------------------

        MensagemParcial = ""

        If txtPedido.Text <> "0" And txtPedido.Text <> "" Then
            If ddlMoeda.SelectedValue = 1 Then
                If CDbl(txtValorCobrado.Text) <> txtValorDocumento.Value Then
                    If CDbl(txtValorDoDocumento.Text) < txtValorDocumento.Value And (txtValorDocumento.Value - CDbl(txtValorDoDocumento.Text) > 0) Then
                        GravaTituloParcial()
                    End If
                End If
            End If

            If ddlMoeda.SelectedValue = 3 Then
                If CDbl(txtValorEmMoeda.Text) < txtValorMoeda.Value And (txtValorMoeda.Value - CDbl(txtValorEmMoeda.Text) > 0) Then
                    GravaTituloParcial()
                End If
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso Not CInt(txtPedido.Text) = "0" Then
            Cliente = DdlEmpresaCliente.SelectedValue
            Campo = Cliente.Split("-")
            objPedido = New [Lib].Negocio.Pedido(Campo(0), Campo(1), txtPedido.Text)
        End If

        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
        If DdlProvisoes.SelectedValue = 1 AndAlso Not lblAgrupar.Text = "AP" Then  'Adiantamento e AdiantamentoXBaixa
            Dim ValorAdiantamento As Decimal = CDec(txtValorDoDocumento.Text)
            Dim ValorAdiantamentoMoeda As Decimal = CDec(txtValorEmMoeda.Text)

            If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento Then
                If Not Adiantamento(objPedido, txtRegistro.Text, ValorAdiantamento, ValorAdiantamentoMoeda) Then
                    MsgBox(Me.Page, "Erro no Adiantamento")
                    Exit Sub
                End If
            End If

            If objCarteira.BaixaAdiantamento Then
                If Trim(txtNumeroAdto.Text) <> "" And Trim(txtNumeroAdto.Text) <> "0" Then
                    If Not AdiantamentoAmortizacao(objPedido, txtRegistro.Text, txtNumeroAdto.Text, ValorAdiantamento, ValorAdiantamentoMoeda) Then
                        MsgBox(Me.Page, "Erro na Baixa de Adiantamento")
                        Exit Sub
                    End If
                End If
            End If
        End If

        'Grupamento de Titulos
        '--------------------------------------------------------
        If lblAgrupar.Text = "AP" Then
            For index = 0 To Session("ssRegistros" & HID.Value).Count - 1
                Sql = " SELECT CR.Registro_Id, CR.Sequencia_Id, CR.Provisao, CR.Carteira, CR.Tributo, CR.Indexador, " & vbCrLf & _
                      "        CR.Moeda, CR.TipoPagto, CR.Situacao, CR.Lote, CR.Movimento, CR.Vencimento, CR.Prorrogacao, " & vbCrLf & _
                      "        CR.DataMoeda, CR.Baixa, CR.UnidadeDeNegocio, CR.Empresa, CR.EndEmpresa, CR.Cliente, " & vbCrLf & _
                      "        CR.EndCliente, CR.BancoCliente, CR.AgenciaCliente, CR.DigitoAgenciaCliente, CR.ContaCliente, " & vbCrLf & _
                      "        CR.DigitoContaCliente, Isnull(CR.TipoContaCliente,'C') AS TipoContaCliente, CR.ContaContabilCliente, " & vbCrLf & _
                      "        CR.EmpresaPagadora, CR.EndEmpresaPagadora, CR.BancoPagador, CR.AgenciaPagadora, CR.DigitoAgenciaPagadora, " & vbCrLf & _
                      "        CR.ContaPagadora, CR.DigitoContaPagadora, CR.ContaContabilPagadora, CR.Cheque, isnull(CR.Slips,'N') AS Slips, " & vbCrLf & _
                      "        CR.Recibo, CR.Aviso, CR.ReciboDeposito, isnull(CR.EmpresaPedido,'') AS EmpresaPedido, isnull(CR.EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(CR.Pedido, 0) AS Pedido, " & vbCrLf & _
                      "        isnull(CR.PedidoFixacao,0) AS PedidoFixacao, isnull(CR.Procuracao,0) AS Procuracao, CR.ValorDoDocumento, CR.Descontos, CR.Deducoes, " & vbCrLf & _
                      "        CR.Juros, CR.Acrescimos, CR.ValorLiquido, ISNULL(CR.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, " & vbCrLf & _
                      "        ISNULL(CR.MoedaDescontos, 0) AS MoedaDescontos, ISNULL(CR.MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(CR.MoedaJuros, 0) AS MoedaJuros, " & vbCrLf & _
                      "        ISNULL(CR.MoedaAcrescimos, 0) AS MoedaAcrescimos, ISNULL(CR.MoedaValorLiquido, 0) AS MoedaValorLiquido, CR.Historico, " & vbCrLf & _
                      "        CR.CodigoDeBarras, CR.CodigoDigitado, CR.Destinatario, CR.EndDestinatario, CR.NomeDoDestinatario, CR.Destinacao, " & vbCrLf & _
                      "        CR.Solicitacao, CR.UsuarioInclusao, CR.UsuarioInclusaoData, CR.UsuarioAlteracao, CR.UsuarioAlteracaoData, " & vbCrLf & _
                      "        CR.UsuarioCancelamento, CR.UsuarioCancelamentoData, isnull(CR.UsuarioLiberacao,'') AS UsuarioLiberacao, CR.UsuarioLiberacaoData, " & vbCrLf & _
                      "        CR.UsuarioBaixa, CR.UsuarioBaixaData, isnull(CR.Grupado,'N') AS Grupado, isnull(CR.RegistroMestre, 0) as RegistroMestre, CR.Observacoes, " & vbCrLf & _
                      "        CR.SituacaoBancaria, ISNULL(CR.NumeroDoCheque,0) AS NumeroDoCheque, " & vbCrLf & _
                      "        isnull(CR.UsuarioLiberacaoBloqueio, '') AS UsuarioLiberacaoBloqueio, CR.UsuarioLiberacaoBloqueioDate, isnull(CR.UsuarioLiberacaoPedido, '') AS UsuarioLiberacaoPedido, " & vbCrLf & _
                      "        CR.UsuarioLiberacaoPedidoDate, isnull(CR.UsuarioLiberacaoCheque, '') AS UsuarioLiberacaoCheque, CR.UsuarioLiberacaoChequeDate, " & vbCrLf & _
                      "        isnull(CR.CarteiraAdto,'') AS CarteiraAdto, isnull(CR.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(CR.ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, " & vbCrLf & _
                      "        ISNULL(NF.TipoDeDocumento,0) AS TipoDeDocumento " & vbCrLf & _
                      "   FROM ContasAReceber CR " & vbCrLf & _
                      "   LEFT JOIN NotaFiscalXTitulo NFxT " & vbCrLf & _
                      "		ON CR.Registro_Id = NFxT.Titulo_Id " & vbCrLf & _
                      "   LEFT JOIN NotasFiscais NF " & vbCrLf & _
                      "		ON NFxT.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                      "	   AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                      "	   AND NFxT.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                      "    AND NFxT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                      "    AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                      "    AND NFxT.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                      "    AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                      "  WHERE CR.Registro_Id = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & " and CR.Situacao not in (2,3,4,5,6,10) " & vbCrLf

                Dim dsFilho As New DataSet
                dsFilho = Banco.ConsultaDataSet(Sql, "ContasAReceberXFilho")

                If Not dsFilho Is Nothing AndAlso dsFilho.Tables(0).Rows.Count > 0 Then
                    For Each drFilho As DataRow In dsFilho.Tables(0).Rows
                        Sql = "DELETE FROM razao" & vbCrLf & _
                              " WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index))
                        SqlArray.Add(Sql)

                        Sql = "DELETE FROM Adiantamentos" & vbCrLf & _
                              " WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index))
                        SqlArray.Add(Sql)

                        Sql = "DELETE FROM AdiantamentosXBaixas" & vbCrLf & _
                              " WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index))
                        SqlArray.Add(Sql)

                        Sql = "UPDATE ContasAReceber" & vbCrLf & _
                              "   SET Provisao       = " & DdlProvisoes.SelectedValue & "," & vbCrLf & _
                              "       TipoPagto      = '" & DdlTiposDePagamentos.SelectedValue & "'," & vbCrLf & _
                              "       Movimento      = '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'," & vbCrLf & _
                              "       DataMoeda = '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'," & vbCrLf & _
                              "       Baixa = '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'," & vbCrLf & _
                              "       Prorrogacao = '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'"

                        If DdlTributos.SelectedIndex > 0 Then
                            Sql &= ", Tributo = '" & DdlTributos.SelectedValue & "'"
                        End If

                        If drFilho("Moeda") = 1 AndAlso Not ddlIndexador.SelectedValue = 99 Then
                            If CDec(drFilho("MoedaValorDoDocumento")) = 0 Then drFilho("MoedaValorDoDocumento") = DolarizaBaixa(drFilho("ValorDoDocumento"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                            If CDec(drFilho("Descontos")) > 0 Then drFilho("MoedaDescontos") = Math.Round(drFilho("Descontos") / (drFilho("ValorDoDocumento") / drFilho("MoedaValorDoDocumento")), 2, MidpointRounding.AwayFromZero)
                            If CDec(drFilho("Deducoes")) > 0 Then drFilho("MoedaDeducoes") = Math.Round(drFilho("Deducoes") / (drFilho("ValorDoDocumento") / drFilho("MoedaValorDoDocumento")), 2, MidpointRounding.AwayFromZero)
                            If CDec(drFilho("Juros")) > 0 Then drFilho("MoedaJuros") = Math.Round(drFilho("Juros") / (drFilho("ValorDoDocumento") / drFilho("MoedaValorDoDocumento")), 2, MidpointRounding.AwayFromZero)
                            If CDec(drFilho("Acrescimos")) > 0 Then drFilho("MoedaAcrescimos") = Math.Round(drFilho("Acrescimos") / (drFilho("ValorDoDocumento") / drFilho("MoedaValorDoDocumento")), 2, MidpointRounding.AwayFromZero)
                            drFilho("MoedaValorLiquido") = drFilho("MoedaValorDoDocumento") + drFilho("MoedaJuros") + drFilho("MoedaAcrescimos") - drFilho("MoedaDescontos") - drFilho("MoedaDeducoes")
                        End If
                        'Valor Documento Moeda
                        Sql &= ", MoedaValorDoDocumento = " & Replace(drFilho("MoedaValorDoDocumento"), ",", ".")
                        'Descontos
                        Sql &= ", MoedaDescontos = " & Replace(drFilho("MoedaDescontos"), ",", ".")
                        'Deducoes
                        Sql &= ", MoedaDeducoes = " & Replace(drFilho("MoedaDeducoes"), ",", ".")
                        'Juros
                        Sql &= ", MoedaJuros = " & Replace(drFilho("MoedaJuros"), ",", ".")
                        'Acrescimos
                        Sql &= ", MoedaAcrescimos = " & Replace(drFilho("MoedaAcrescimos"), ",", ".")
                        'Liquido
                        Sql &= ", MoedaValorLiquido = " & Replace(drFilho("MoedaValorLiquido"), ",", ".")

                        Sql &= ", RegistroMestre = " & txtRegistro.Text
                        Sql &= ", Grupado = 'S'"

                        Cliente = DdlEmpresaPagadora.SelectedValue
                        Campo = Cliente.Split("-")

                        If Campo(0) <> "" Then
                            Sql &= ", EmpresaPagadora = '" & Campo(0) & "'"                           'EmpresaPagadora
                            Sql &= ", EndEmpresaPagadora = " & Campo(1)                               'Endereco Empresa Pagadora
                        Else
                            Sql &= ", EmpresaPagadora = ''"
                            Sql &= ", EndEmpresaPagadora = 0"
                        End If

                        If DdlBancoRecebedor.SelectedIndex <> 0 Then
                            Sql &= ", BancoPagador = '" & DdlBancoRecebedor.SelectedValue & "'"          'Banco
                            Cliente = DdlContaRecebedora.SelectedValue
                            Campo = Cliente.Split("-")
                            Sql &= ", AgenciaPagadora = '" & Campo(0) & "'"                           'Agencia Cliente
                            Sql &= ", DigitoAgenciaPagadora = '" & Campo(1) & "'"                           'Digito Agencia Cliente
                            Sql &= ", ContaPagadora = '" & Campo(2) & "'"                           'Conta Cliente
                            Sql &= ", DigitoContaPagadora = '" & Campo(3) & "'"                           'Digito Conta Cliente
                            Sql &= ", ContaContabilPagadora = '" & Campo(4) & "'"                           'Conta Contabil
                        End If

                        Sql &= ", UsuarioAlteracao = '" & Session("ssNomeUsuario") & "'"                'Usuario que Incluiu
                        Sql &= ", UsuarioAlteracaoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"  'Data da Inclusao
                        Sql &= ", ContratoBancario = '" & Funcoes.EliminarCaracteresEspeciais(txtContratoBanco.Text) & "'" 'Contrato Bancário

                        Sql &= " WHERE Registro_ID = " & CStr(Session("ssRegistros" & HID.Value).Item(index))

                        SqlArray.Add(Sql)

                        If DdlProvisoes.SelectedValue = 1 Then
                            ConsultaCarteiras(drFilho("Carteira"))               'Ler Contas Contabeis da Carteira
                            ' Grava Razao Debito
                            '------------------------------------------
                            Registro = drFilho("Registro_Id")
                            Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                            Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente

                            Raz_Conta = ContaClientes                            'Conta sem tributo
                            Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                            If objPlaConta.TemCliente Then
                                Raz_Cliente = drFilho("Cliente")                'Cliente
                                Raz_EndCliente = drFilho("EndCliente")
                            Else
                                Raz_Cliente = ""                                'Cliente
                                Raz_EndCliente = 0                              'Endereco do Cliente
                            End If

                            If drFilho("Tributo") <> "" Then
                                ConsultaTributos(drFilho("Tributo"))      'Consulta Contas Contabeis dos Tributos

                                Raz_Conta = ContaClientes                       'Conta com tributo
                                Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                                If objPlaContaTributo.TemCliente Then

                                    Raz_Cliente = drFilho("Cliente")                'Cliente
                                    Raz_EndCliente = drFilho("EndCliente")
                                Else
                                    Raz_Cliente = ""                                'Cliente
                                    Raz_EndCliente = 0                              'Endereco do Cliente
                                End If
                            End If

                            Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio

                            Raz_ValorOficial = drFilho("ValorDoDocumento")       'ValorDoDocumento

                            If drFilho("MoedaValorDoDocumento") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                Raz_ValorMoeda = drFilho("MoedaValorDoDocumento")
                            Else
                                Raz_ValorMoeda = Replace(DolarizaBaixa(drFilho("ValorDoDocumento"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                            End If

                            Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                 'Historico
                            Raz_DebitoCredito = "C"                         'Debito/Credito

                            LanctosContabeis()

                            '-------Descontos-----------
                            If drFilho("Descontos") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = ContaDescontos                      'Grupo de Contas
                                Raz_Cliente = ""                                'Cliente
                                Raz_EndCliente = 0                              'Endereco do Cliente
                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio

                                Raz_ValorOficial = drFilho("Descontos")                  'ValorDoDocumento

                                If drFilho("MoedaDescontos") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaDescontos")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(drFilho("Descontos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                'Historico
                                Raz_DebitoCredito = "D"                         'Debito/Credito

                                LanctosContabeis()
                            End If

                            '----Deducoes--------------
                            If drFilho("Deducoes") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = ContaDeducoes                       'Grupo de Contas
                                Raz_Cliente = ""                                'Cliente
                                Raz_EndCliente = 0                              'Endereco do Cliente
                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio
                                Raz_ValorOficial = drFilho("Deducoes")               'ValorDoDocumento

                                If drFilho("MoedaDeducoes") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaDeducoes")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(drFilho("Deducoes"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                 'Historico
                                Raz_DebitoCredito = "D"                         'Debito/Credito
                                LanctosContabeis()
                            End If

                            '----Juros--------------
                            If drFilho("Juros") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = ContaJuros                          'Grupo de Contas
                                Raz_Cliente = ""                                'Cliente
                                Raz_EndCliente = 0                              'Endereco do Cliente
                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio

                                Raz_ValorOficial = drFilho("Juros")                     'ValorDoDocumento

                                If drFilho("MoedaJuros") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaJuros")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(drFilho("Juros"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                 'Historico
                                Raz_DebitoCredito = "C"                         'Debito/Credito
                                LanctosContabeis()
                            End If

                            '----Acrescimos--------------
                            If drFilho("Acrescimos") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = ContaAcrescimos                     'Grupo de Contas
                                Raz_Cliente = ""                                'Cliente
                                Raz_EndCliente = 0                              'Endereco do Cliente
                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio
                                Raz_ValorOficial = drFilho("Acrescimos")                    'ValorDoDocumento

                                If drFilho("MoedaAcrescimos") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaAcrescimos")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(drFilho("Acrescimos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                'Historico
                                Raz_DebitoCredito = "C"                         'Debito/Credito

                                LanctosContabeis()
                            End If
                        End If
                    Next
                End If
            Next
        End If
        '----Fim Contabilização de Baixa por agrupamentos
        '--------------------------------------------------------
        ConsultaCarteiras(ddlCarteiras.SelectedValue)

        '-Gravação no Razão Contabil------------------
        If DdlProvisoes.SelectedValue = 1 And lblAgrupar.Text = "" Then
            ' Grava Razao Credito

            Registro = CInt(txtRegistro.Text)
            Cliente = DdlEmpresaCliente.SelectedValue
            Campo = Cliente.Split("-")
            Raz_Empresa = Campo(0)                  'EmpresaCliente
            Raz_EndEmpresa = Campo(1)               'Endereco Empresa Cliente

            Raz_Conta = ContaClientes                       'Conta sem tributo
            Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
            If objPlaConta.TemCliente Then

                Cliente = txtCodigoCliente.Value
                Campo = Cliente.Split("-")
                Raz_Cliente = Campo(0)                              'Cliente
                Raz_EndCliente = Campo(1)                           'Endereco do Cliente
            Else
                Raz_Cliente = ""                                'Cliente
                Raz_EndCliente = 0                              'Endereco do Cliente
            End If

            If DdlTributos.Text <> "" Then
                ConsultaTributos(DdlTributos.SelectedValue)      'Consulta Contas Contabeis dos Tributos

                Raz_Conta = ContaClientes                       'Conta com tributo
                Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaContaTributo.TemCliente Then

                    Cliente = txtCodigoCliente.Value
                    Campo = Cliente.Split("-")
                    Raz_Cliente = Campo(0)                              'Cliente
                    Raz_EndCliente = Campo(1)                           'Endereco do Cliente
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If
            End If

            Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue
            Raz_ValorOficial = Replace(txtValorDoDocumento.Text, ".", "")

            If CDbl(txtValorEmMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                Raz_ValorMoeda = Replace(txtValorEmMoeda.Text, ".", "")
            Else
                Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
            End If

            Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)

            Raz_DebitoCredito = "C"

            LanctosContabeis()

            '-------Descontos-----------
            If txtDescontos.Text <> "" And txtDescontos.Text > 0 Then
                Registro = CInt(txtRegistro.Text)
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Empresa = Campo(0)                  'EmpresaCliente
                Raz_EndEmpresa = Campo(1)               'Endereco Empresa Cliente

                Raz_Conta = ContaDescontos              'Conta de Descontos Obtidos
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtDescontos.Text, ".", "")  'Valor Contabil

                If CDbl(txtDescontosMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtDescontosMoeda.Text, ".", "")
                Else
                    Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                    Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                    Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)     'Históricos
                Raz_DebitoCredito = "D"

                LanctosContabeis()
            End If
            '----Deducoes--------------
            If txtDeducoes.Text <> "" And txtDeducoes.Text > 0 Then
                Registro = CInt(txtRegistro.Text)
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Empresa = Campo(0)                  'EmpresaCliente
                Raz_EndEmpresa = Campo(1)               'Endereco Empresa Cliente

                Raz_Conta = ContaDeducoes               'Conta de Descontos Obtidos
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtDeducoes.Text, ".", "")  'Valor Contabil

                If CDbl(txtDeducoesMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtDeducoesMoeda.Text, ".", "")
                Else
                    Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                    Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                    Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)     'Históricos
                Raz_DebitoCredito = "D"

                LanctosContabeis()
            End If
            '----Juros--------------
            If txtJuros.Text <> "" And txtJuros.Text > 0 Then
                Registro = CInt(txtRegistro.Text)
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Empresa = Campo(0)                  'EmpresaCliente
                Raz_EndEmpresa = Campo(1)               'Endereco Empresa Cliente

                Raz_Conta = ContaJuros                  'Conta de Descontos Obtidos
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtJuros.Text, ".", "")  'Valor Contabil

                If CDbl(txtJurosMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtJurosMoeda.Text, ".", "")
                Else
                    Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                    Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                    Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
                End If

                'Raz_ValorMoeda = Replace(MoedaJuros, ".", "")  'Valor Contabil

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)     'Históricos
                Raz_DebitoCredito = "C"

                LanctosContabeis()
            End If
            '----Acrescimos--------------
            If txtAcrescimos.Text <> "" And txtAcrescimos.Text > 0 Then
                Registro = CInt(txtRegistro.Text)
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Empresa = Campo(0)                  'EmpresaCliente
                Raz_EndEmpresa = Campo(1)               'Endereco Empresa Cliente

                Raz_Conta = ContaAcrescimos             'Conta de Descontos Obtidos
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo
                Raz_ValorOficial = Replace(txtAcrescimos.Text, ".", "")  'Valor Contabil

                If CDbl(txtAcrescimosMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtAcrescimosMoeda.Text, ".", "")
                Else
                    Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                    Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                    Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
                End If

                'Raz_ValorMoeda = Replace(MoedaAcrescimos, ".", "")  'Valor Contabil
                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)      'Históricos
                Raz_DebitoCredito = "C"

                LanctosContabeis()
            End If
        End If
        '-----------------------
        'Gravar Debito
        '-----------------------
        If DdlProvisoes.SelectedValue = 1 And (lblAgrupar.Text = "" Or lblAgrupar.Text = "AP") Then
            Registro = CInt(txtRegistro.Text)
            Cliente = DdlEmpresaPagadora.SelectedValue
            Campo = Cliente.Split("-")
            Raz_Empresa = Campo(0)                                  'Empresa Pagadora
            Raz_EndEmpresa = Campo(1)                               'Endereco Empresa Pagadora

            Cliente = DdlContaRecebedora.SelectedValue
            Campo = Cliente.Split("-")
            Raz_Conta = Campo(4)                                        'Conta Contabil

            If Len(Raz_Conta) = 7 Then
                Cliente = txtCodigoCliente.Value
                Campo = Cliente.Split("-")
                Raz_Cliente = Campo(0)                                  'Cliente
                Raz_EndCliente = Campo(1)
            Else
                Raz_Cliente = ""                                        'Cliente
                Raz_EndCliente = 0                                      'Endereco do Cliente
            End If

            Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'estava fixo "01"  antes 19/06/12     'Pegar Unidadde de Negocio da EmpresaPagadora

            Raz_ValorOficial = Replace(txtValorCobrado.Text, ".", "")

            If CDbl(txtValorCobradoMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                Raz_ValorMoeda = Replace(txtValorCobradoMoeda.Text, ".", "")
            Else
                Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
            End If

            Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)
            Raz_DebitoCredito = "D"
            LanctosContabeis()
            '-------------------------------------------
            'Transferencias Financeiras
            '-------------------------------------------
            Cliente = DdlEmpresaCliente.SelectedValue       'Empresa do Cliente Empresa de Credito
            Campo = Cliente.Split("-")

            Sql = "SELECT EmpresaDebito,EnderecoDebito,EmpresaCredito,EnderecoCredito,EmpresaContabil,EnderecoContabil,ContaContabil," & vbCrLf & _
                  "       ClienteContabil,EndClienteContabil,case when DebitoCredito='D'then 'C' else 'D' end DebitoCredito " & vbCrLf & _
                  "  FROM TransferenciasFinanceiras " & vbCrLf & _
                  " WHERE EmpresaDebito='" & Campo(0) & "' " & vbCrLf & _
                  "   AND EnderecoDebito = " & Campo(1) & vbCrLf

            Cliente = DdlEmpresaPagadora.SelectedValue   'Empresa Recebedora  Empresa de Debito
            Campo = Cliente.Split("-")
            Sql &= " AND EmpresaCredito = '" & Campo(0) & "'" & vbCrLf & _
                   " AND EnderecoCredito = " & Campo(1)

            For Each DrT As DataRow In Banco.ConsultaDataSet(Sql, "Transferencias").Tables(0).Rows
                Raz_Empresa = DrT("EmpresaContabil")                'EmpresaCliente
                Raz_EndEmpresa = DrT("EnderecoContabil")            'Endereco Empresa Cliente
                Raz_Conta = DrT("ContaContabil")                    'Grupo de Contas
                Raz_Cliente = DrT("ClienteContabil")                'Cliente
                Raz_EndCliente = DrT("EndClienteContabil")          'Endereco do Cliente
                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue

                Raz_ValorOficial = CDec(txtValorDoDocumento.Text) + CDec(txtJuros.Text) + CDec(txtAcrescimos.Text) - CDec(txtDescontos.Text) - CDec(txtDeducoes.Text)

                If ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = CDec(txtValorEmMoeda.Text) + CDec(txtJurosMoeda.Text) + CDec(txtAcrescimosMoeda.Text) - CDec(txtDescontosMoeda.Text) - CDec(txtDeducoesMoeda.Text)
                Else
                    Valor = Replace(DolarizaBaixa(Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                    Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                    Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)               'Historico
                Raz_DebitoCredito = DrT("DebitoCredito")            'Debito/Credito
                LanctosContabeis()
            Next
        End If

        objCarteira = New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        If Not Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
            Dim objDestinoContabil() As String = Session("objDestinoContabil" & HID.Value).ToString.Split("-")

            Sql = " INSERT INTO TitulosXDesdobrarFornecedor (Registro_Id, Cliente, EndCliente, Pedido, Carteira) " & vbCrLf & _
                  " VALUES (" & Registro & "," & vbCrLf & _
                  "        '" & objDestinoContabil(0) & "'," & vbCrLf & _
                  "         " & objDestinoContabil(1) & "," & vbCrLf & _
                  "         " & 0 & "," & vbCrLf & _
                  "        '" & objDestinoContabil(2) & "')" & vbCrLf

            SqlArray.Add(Sql)
        End If

        If Banco.GravaBanco(SqlArray) Then
            GridConsultaTitulos.DataSource = Nothing
            GridConsultaTitulos.DataBind()
            MsgBox(Me.Page, "Registro <" & Registro & "> Gravado/Atualizado com sucesso." & MensagemParcial)
            txtRegistro.Text = Registro
            If chkEmitirRecibo.Checked Then
                EmitirRecibo()
            End If
            Limpar(True)
        End If
    End Sub

    Sub GravaTituloParcial()

        Dim Valor As String = ""
        Dim Opcao As String = txtRegistro.Text
        Dim Registro As Integer = 0
        '''''''''''''''''''''''''''''''
        'Gera sequencia de titulos Parcial'''
        ''''''''''''''''''''''''''''''
        Sql = "exec sp_Numerador '" & Session("ssNomeServidor").ToUpper() & "',0,1"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
            Registro = Dr("Sequencia")
        Next

        Data = txtMovimento.Text

        Sql = "INSERT INTO ContasAReceber " & vbCrLf & _
              "       (Registro_Id" & vbCrLf & _
              "       ,Sequencia_Id" & vbCrLf & _
              "       ,Provisao" & vbCrLf & _
              "       ,Carteira" & vbCrLf & _
              "       ,Tributo" & vbCrLf & _
              "       ,Indexador" & vbCrLf & _
              "       ,Moeda" & vbCrLf & _
              "       ,TipoPagto" & vbCrLf & _
              "       ,Situacao" & vbCrLf & _
              "       ,Lote" & vbCrLf & _
              "       ,Movimento" & vbCrLf & _
              "       ,Vencimento" & vbCrLf & _
              "       ,Prorrogacao" & vbCrLf & _
              "       ,DataMoeda" & vbCrLf & _
              "       ,Baixa" & vbCrLf & _
              "       ,UnidadeDeNegocio" & vbCrLf & _
              "       ,Empresa" & vbCrLf & _
              "       ,EndEmpresa" & vbCrLf & _
              "       ,Cliente" & vbCrLf & _
              "       ,EndCliente" & vbCrLf & _
              "       ,BancoCliente" & vbCrLf & _
              "       ,AgenciaCliente" & vbCrLf & _
              "       ,DigitoAgenciaCliente" & vbCrLf & _
              "       ,ContaCliente" & vbCrLf & _
              "       ,DigitoContaCliente" & vbCrLf & _
              "       ,ContaContabilCliente" & vbCrLf & _
              "       ,EmpresaPagadora" & vbCrLf & _
              "       ,EndEmpresaPagadora" & vbCrLf & _
              "       ,BancoPagador" & vbCrLf & _
              "       ,AgenciaPagadora" & vbCrLf & _
              "       ,DigitoAgenciaPagadora" & vbCrLf & _
              "       ,ContaPagadora" & vbCrLf & _
              "       ,DigitoContaPagadora" & vbCrLf & _
              "       ,ContaContabilPagadora" & vbCrLf & _
              "       ,Cheque" & vbCrLf & _
              "       ,Slips" & vbCrLf & _
              "       ,Recibo" & vbCrLf & _
              "       ,Aviso" & vbCrLf & _
              "       ,ReciboDeposito" & vbCrLf & _
              "       ,EmpresaPedido" & vbCrLf & _
              "       ,EndEmpresaPedido" & vbCrLf & _
              "       ,Pedido" & vbCrLf & _
              "       ,PedidoFixacao" & vbCrLf & _
              "       ,ValorDoDocumento" & vbCrLf & _
              "       ,Descontos" & vbCrLf & _
              "       ,Deducoes" & vbCrLf & _
              "       ,Juros" & vbCrLf & _
              "       ,Acrescimos" & vbCrLf & _
              "       ,ValorLiquido" & vbCrLf & _
              "       ,MoedaValorDoDocumento" & vbCrLf & _
              "       ,MoedaDescontos" & vbCrLf & _
              "       ,MoedaDeducoes" & vbCrLf & _
              "       ,MoedaJuros" & vbCrLf & _
              "       ,MoedaAcrescimos" & vbCrLf & _
              "       ,MoedaValorLiquido" & vbCrLf & _
              "       ,Historico" & vbCrLf & _
              "       ,CodigoDeBarras" & vbCrLf & _
              "       ,CodigoDigitado" & vbCrLf & _
              "       ,Destinatario" & vbCrLf & _
              "       ,EndDestinatario" & vbCrLf & _
              "       ,solicitacao" & vbCrLf & _
              "       ,UsuarioInclusao" & vbCrLf & _
              "       ,UsuarioInclusaoData" & vbCrLf & _
              "       ,UsuarioBaixa" & vbCrLf & _
              "       ,UsuarioBaixaData" & vbCrLf & _
              "       ,Grupado" & vbCrLf & _
              "       ,Observacoes)" & vbCrLf & _
              " VALUES( " & vbCrLf

        Sql &= Registro                                             'Registro
        Sql &= ", 0"                                                'Sequencia
        Sql &= ", 2"
        Sql &= ", '" & ddlCarteiras.SelectedValue & "'"
        Sql &= ", ''"                                               'Tributo
        Sql &= ", " & ddlIndexador.SelectedValue                    'Indexador
        Sql &= ", " & ddlMoeda.SelectedValue                        'Moeda
        Sql &= ", 0"                                                'Tipo de Pagamento
        Sql &= ", 1"                                                'Situacao
        Sql &= ", 70"                                               'Lote

        'Caso o título principal (que teve seu valor alterado) esteja na tabela NotaFiscalXTitulo então seus dados são carregados.
        'Esses dados serão utilizados nesse momento para validação e mais abaixo para inserção deste registro parcial na tabela NotafiscalXTitulo

        Dim dsNFxTitulo = CarregarNotaFiscalXTitulo()

        If Not dsNFxTitulo.Tables Is Nothing AndAlso dsNFxTitulo.Tables(0).Rows.Count > 0 Then
            Sql &= ", '" & CDate(hdnMovimentoOriginal.Value).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(hdnProrrogacaoOriginal.Value).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(hdnProrrogacaoOriginal.Value).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(hdnProrrogacaoOriginal.Value).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(hdnMovimentoOriginal.Value).ToString("yyyy/MM/dd") & "'" & vbCrLf
        Else
            Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
        End If

        Sql &= ", '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'"


        Dim Cliente() As String = Nothing
        'Dim Campo() As String = Nothing

        Cliente = DdlEmpresaCliente.SelectedValue.Split("-")
        'Cliente = Cliente.Split("-")

        If Cliente(0) <> "" Then
            Sql &= ", '" & campo(0) & "'"                           'EmpresaCliente
            Sql &= ", " & campo(1)                                  'Endereco Empresa Cliente
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If


        Cliente = txtCodigoCliente.Value.Split("-")
        'campo = Cliente.Split("-")

        If Cliente(0) <> "" Then
            Sql &= ", '" & Cliente(0) & "'"                           'Cliente
            Sql &= ", " & Cliente(1)                                  'Endereco Cliente
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If

        '-Dados Bancarios Cliente----------

        Sql &= ", 0"                        'Banco Cliente
        Sql &= ", ''"                       'Agencia do Destinatario
        Sql &= ", ''"                       'Digito da Agencia do Destinatário
        Sql &= ", ''"                       'Conta Corrente do Destinatário
        Sql &= ", ''"                       'Digito da Conta Corrente do Destinatário

        Sql &= ", '" & Carteira(ddlCarteiras.SelectedValue) & "'"  'Grupo de Conta Corrente do Fornecedor

        'Empresa Pagadora------------------

        Sql &= ", ''"                                           'Empresa Recebedora
        Sql &= ", 0"                                            'Endereco Empresa Recebedora
        Sql &= ", 0"                                            'Banco Recebedor
        Sql &= ", ''"                                           'Agencia Recebedora
        Sql &= ", ''"                                           'Digito Agencia Recebedora
        Sql &= ", ''"                                           'Conta Corrente Recebedora
        Sql &= ", ''"                                           'Digito Conta Recebedora
        Sql &= ", ''"                                           'Digito Conta Recebedora

        '-----------------------------------------------

        Sql &= ", 'N'"                                              'Emite Cheque
        Sql &= ", 'N'"                                              'Emite Slips
        Sql &= ", 'N'"                                              'Emite Recibo
        Sql &= ", 'N'"                                              'Emite Aviso
        Sql &= ", 'N'"                                              'Emite Recibo De Deposito

        If String.IsNullOrWhiteSpace(txtPedido.Text) OrElse CInt(txtPedido.Text) = "0" Then
            Sql &= ", NULL"
            Sql &= ", NULL"
            Sql &= ", NULL"
            Sql &= ", NULL"
        Else
            Cliente = DdlEmpresaCliente.SelectedValue.Split("-")
            'campo = Cliente.Split("-")
            Sql &= ", '" & Cliente(0) & "'"
            Sql &= ", " & Cliente(1)
            Sql &= ", " & txtPedido.Text
            If txtPedidoFixacao.Value.ToString.Length = 0 OrElse txtPedidoFixacao.Value = 0 Then  'PedidoFixacao
                Sql &= ", NULL"
            Else
                Sql &= ", " & txtPedidoFixacao.Value
            End If
        End If

        '------------------------------------
        If txtValorDoDocumento.Text = "" Then                       'Valor do Documento
            Sql &= ", 0"
        Else
            Valor = txtValorDocumento.Value - txtValorDoDocumento.Text
            Valor = Replace(Valor, ".", "")
            Sql &= ", " & Replace(Valor, ",", ".")
        End If

        Sql &= ", 0" 'Descontos
        Sql &= ", 0" 'Deducoes
        Sql &= ", 0" 'Juros
        Sql &= ", 0" 'Acrescimos

        Sql &= ", " & Replace(Valor, ",", ".")

        '-------------------------------------
        '---Valores em Dolar------------------

        If txtValorEmMoeda.Text = "" Then                       'Valor do Documento
            Sql &= ", 0"
        ElseIf ddlIndexador.SelectedValue = 99 Then
            Sql &= ", 0"
        Else
            Dim Sql2 As String

            Sql2 = "SELECT ISNULL(IndiceFixado,0) AS IndiceFixado " & vbCrLf & vbCrLf & _
                   "  FROM Pedidos " & vbCrLf & _
                   " WHERE Pedido_id = " & txtPedido.Text & vbCrLf
            Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql2, "Pedidos")

            If ddlMoeda.SelectedValue = 1 AndAlso txtPedido.Text > 0 AndAlso dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                Valor = Math.Round((CDbl(txtValorDocumento.Value) - CDbl(txtValorDoDocumento.Text)) / dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                Sql &= ", " & Replace(Valor, ",", ".")
            Else
                Valor = txtValorMoeda.Value - txtValorEmMoeda.Text
                Valor = Replace(Valor, ".", "")
                Sql &= ", " & Replace(Valor, ",", ".")
            End If
        End If

        Sql &= ", 0" 'Descontos
        Sql &= ", 0" 'Deducoes
        Sql &= ", 0" 'Juros
        Sql &= ", 0" 'Acrescimos

        If ddlIndexador.SelectedValue = 99 Then
            Sql &= ", 0"
        Else
            Sql &= ", " & Replace(Valor, ",", ".")
        End If

        '-------------------------------------

        If txtPedido.Text <> "0" Then
            If ddlMoeda.SelectedValue = 1 Then
                txtHistorico.Text = PesoPago(txtPedido.Text, (txtValorDocumento.Value - txtValorDoDocumento.Text), txtHistorico.Text)
            Else
                txtHistorico.Text = PesoPago(txtPedido.Text, (txtValorMoeda.Value - txtValorEmMoeda.Text), txtHistorico.Text)
            End If
        End If

        Sql &= ", '" & txtHistorico.Text & "'"                  'Historico
        Sql &= ", ''"                                           'Codigo De Barras
        Sql &= ", 'N'"                                          'Digitado

        Cliente = txtCodigoCliente.Value.Split("-")
        'campo = Cliente.Split("-")

        If Cliente(0) <> "" Then
            Sql &= ", '" & Cliente(0) & "'"                          'Cliente 
            Sql &= ", " & Cliente(1)                                 'Endereco Cliente 
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If

        '--------------------------------------------------------
        Sql &= ", 0"                                                                'Solicitacao

        Sql &= ", '" & Session("ssNomeUsuario") & "'"                                   'Usuario que Incluiu
        Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                     'Data da Inclusao

        Sql &= ", '" & Session("ssNomeUsuario") & "'"                                   'Usuario que Baixou
        Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                     'Data da Baixa

        If Session("ssGrupado" & HID.Value) = "S" Then
            Sql &= ", 'S'"                                                          'Registro Grupado
            Sql &= ", '" & Session("ssObservacoes" & HID.Value) & "')"      'Observaçoes de Agrupamento
        Else
            Sql &= ", 'N'"                                                          'Registro Grupado
            Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text) & " ')"                              'Observacoes
        End If

        SqlArray.Add(Sql)

        Dim Empresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
        Dim Fornecedor() As String = txtCodigoCliente.Value.Split("-")


        'relaciona a nota ao titulo quando o título desmembrado estiver associado a uma nota
        If Not dsNFxTitulo.Tables Is Nothing AndAlso dsNFxTitulo.Tables(0).Rows.Count > 0 Then
            Sql = " INSERT INTO NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id) " & vbCrLf & _
                  " VALUES('" & Empresa(0) & "'," & Empresa(1) & ",'" & Cliente(0) & "'," & Cliente(1) & ",'" & dsNFxTitulo.Tables(0).Rows(0).Item("EntradaSaida_Id") & "','" & dsNFxTitulo.Tables(0).Rows(0).Item("Serie_Id") & "'," & dsNFxTitulo.Tables(0).Rows(0).Item("Nota_Id") & "," & Registro & ")"
            SqlArray.Add(Sql)
        End If
        '**************************************

        'Adiciona o novo título parcial à tabela FaturaDeFreteXTitulo caso o mesmo seja oriundo de um título de fatura
        If CInt(txtCodigoFaturaDeFrete.Text) > 0 Then
            Dim FFTemp As New FaturaDeFrete() With {.CodigoEmpresa = Empresa(0),
                                                    .EnderecoEmpresa = Empresa(1),
                                                    .CodigoConveniado = Cliente(0),
                                                    .EnderecoConveniado = CInt(Cliente(1).ToString),
                                                    .CodigoFatura = CInt(txtCodigoFaturaDeFrete.Text)}
            Dim FFxT As New FaturaDeFretexTitulo(FFTemp)
            FFxT.CodigoTitulo = Registro
            FFxT.IUD = "I"
            FFxT.SalvarSql(SqlArray)
            '********************************
        End If
        MensagemParcial = " E Registro Parcial de Número <" & Registro & ">..."
    End Sub

    Private Function CarregarNotaFiscalXTitulo() As DataSet
        'verifica se o título desmembrado esta associado a uma nota
        Dim sqlNotaXTiulo = " SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id " & vbCrLf & _
                            "   FROM NotaFiscalXTitulo    " & vbCrLf & _
                            "  WHERE Titulo_Id = " & txtRegistro.Text

        Return Banco.ConsultaDataSet(sqlNotaXTiulo, "NotaXTitulo")
    End Function


    Private Function Adiantamento(ByVal Pedido As [Lib].Negocio.Pedido, ByVal RegistroTitulo As String, ByVal ValorAdtoOficial As Decimal, ByVal ValorAdtoMoeda As Decimal) As Boolean
        Dim RegistroAdiantamento As Integer = 0

        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        Sql = "exec sp_Numerador '" & campo(0) & "', " & campo(1) & ", 15"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
            RegistroAdiantamento = Dr("Sequencia")
        Next

        If RegistroAdiantamento = 0 Then
            MensagemParcial = "Registro <" & RegistroAdiantamento & "> Numerador 15 Adiantamentos não encontrado..."
            Return False
        End If


        Sql = " INSERT INTO Adiantamentos" & _
                " (Empresa_ID " & _
                ", EndEmpresa_ID" & _
                ", Cliente_ID" & _
                ", EndCliente_ID" & _
                ", Adiantamento_Id" & _
                ", RegistroPedido" & _
                ", Titulo" & _
                ", Recibo" & _
                ", Safra" & _
                ", Movimento" & _
                ", Vencimento" & _
                ", ValorOficial" & _
                ", ValorMoeda" & _
                ", Indexador" & _
                ", Moeda" & _
                ", Taxa)" & _
                " VALUES('" & campo(0) & "'" & _
                "," & campo(1)

        campo = txtCodigoCliente.Value.Split("-")
        Sql &= ", '" & campo(0) & "'"                           'Cliente
        Sql &= ", " & campo(1)                                  'Endereco Cliente

        Sql &= ", " & RegistroAdiantamento                      'Numerador de Adiantamentos

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then 'Numero do Pedido
            Sql &= ", NULL"
        Else
            Sql &= ", " & Pedido.Codigo
        End If

        Sql &= ", " & RegistroTitulo                            'Titulo

        Sql &= ", 0"                                            'Recibo

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then 'Safra
            Sql &= ", NULL"
        Else
            Sql &= ", '" & Pedido.CodigoSafra & "'"
        End If


        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'"
        Sql &= ", '" & CDate(txtVencimentoAdto.Text).ToString("yyyy/MM/dd") & "'"

        Sql &= ", " & Str(ValorAdtoOficial)
        Sql &= ", " & Str(ValorAdtoMoeda)

        Sql &= ", " & ddlIndexador.SelectedValue                 'Indexador
        Sql &= ", " & ddlMoeda.SelectedValue                     'Moeda
        Sql &= ", 0)"  'Taxa

        SqlArray.Add(Sql)

        Sql = " Update contasAReceber set" & _
              "  Adiantamento = " & RegistroAdiantamento & _
              " Where registro_id = " & RegistroTitulo

        SqlArray.Add(Sql)

        Return True
    End Function

    Private Function AdiantamentoAmortizacao(ByVal Pedido As [Lib].Negocio.Pedido, ByVal RegistroTitulo As String, ByVal NumeroAdto As String, ByVal ValorBXAdtoOficial As Decimal, ByVal ValorBXAdtoMoeda As Decimal) As Boolean
        Dim objAdiantamento As New [Lib].Negocio.Adiantamento()
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")
        objAdiantamento.CodigoEmpresa = campo(0)
        objAdiantamento.EndEmpresa = campo(1)

        Cliente = txtCodigoCliente.Value
        campo = Cliente.Split("-")
        objAdiantamento.CodigoCliente = campo(0)
        objAdiantamento.EndCliente = campo(1)
        objAdiantamento.Codigo = CInt(NumeroAdto)

        objAdiantamento = New [Lib].Negocio.Adiantamento(0, CInt(NumeroAdto), DdlEmpresaCliente.SelectedValue.ToString.Split("-")(0), DdlEmpresaCliente.SelectedValue.ToString.Split("-")(1))

        If objAdiantamento.Codigo = 0 Then
            'Mensagem = "Adiantamento informado não foi encontrado, verifique a lista."
            Return False
        End If

        Dim Sequencia As Integer

        Sqla = "  SELECT ISNULL(MAX(Sequencia_Id), 0) + 1 AS Sequencia  " & vbCrLf & _
               "    FROM AdiantamentosXBaixas" & vbCrLf & _
               "   WHERE Empresa_Id      ='" & objAdiantamento.CodigoEmpresa & "'" & vbCrLf & _
               "     AND EndEmpresa_Id   ='" & objAdiantamento.EndEmpresa & "'" & vbCrLf & _
               "     AND Adiantamento_Id = " & NumeroAdto & vbCrLf


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
            Sequencia = Dr("Sequencia")
        Next

        Sql = " INSERT INTO AdiantamentosXBaixas" & _
                    " (Empresa_ID " & _
                    ", EndEmpresa_ID" & _
                    ", Cliente_ID" & _
                    ", EndCliente_ID" & _
                    ", Adiantamento_Id" & _
                    ", Sequencia_Id" & _
                    ", RegistroPedido" & _
                    ", Titulo" & _
                    ", ValorOficial" & _
                    ", ValorMoeda" & _
                    ", VariacaoOficial" & _
                    ", VariacaoMoeda" & _
                    ", DataBaixa)" & _
                    " VALUES('" & objAdiantamento.CodigoEmpresa & "'" & _
                    "," & objAdiantamento.EndEmpresa


        Sql &= ", '" & objAdiantamento.CodigoCliente & "'"                          'Cliente
        Sql &= ", " & objAdiantamento.EndCliente                                    'Endereco Cliente

        Sql &= ", " & NumeroAdto                                                    'Numero do Adiantamentos
        Sql &= ", " & Sequencia                                                     'Sequencia

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then                          'Numero do Pedido
            Sql &= ", NULL" & vbCrLf
        Else
            Sql &= ", " & Pedido.Codigo & vbCrLf
        End If

        Sql &= ", " & RegistroTitulo                                                'Titulo


        Sql &= ", " & Str(ValorBXAdtoOficial)
        Sql &= ", " & Str(ValorBXAdtoMoeda)


        Sql &= ", 0"
        Sql &= ", 0"
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "')"

        SqlArray.Add(Sql)

        Return True
    End Function

    'Function AdiantamentoAmortizacao()
    '    Dim cotacao As Double = 0
    '    Dim TemAdiantamento As String = ""
    '    Dim Sequencia As Integer

    '    Cliente = DdlEmpresaCliente.SelectedValue
    '    campo = Cliente.Split("-")

    '    If CInt(txtNumeroAdto.Text) > 0 Then
    '        Sqla = "  SELECT  isnull(MAX(Sequencia_Id), 0) + 1 AS Sequencia  "
    '        Sqla &= " FROM AdiantamentosXBaixas"
    '        Sqla &= " WHERE Empresa_Id = '" & campo(0) & "'"
    '        Sqla &= "   And EndEmpresa_Id = '" & campo(1) & "'"
    '        Sqla &= "   And Adiantamento_Id = " & txtNumeroAdto.Text

    '        Cliente = txtCodigoCliente.Value
    '        campo = Cliente.Split("-")

    '        Sqla &= " And Cliente_Id = '" & campo(0) & "'"
    '        Sqla &= " And EndCliente_Id = '" & campo(1) & "'"

    '        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
    '            Sequencia = Dr("Sequencia")
    '        Next

    '        Sqla = "  SELECT  * "
    '        Sqla &= " FROM ComprasXProdutos"
    '        Sqla &= " WHERE Produto_Id = '" & ddlCarteiras.SelectedValue & "'"

    '        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
    '            Session("ContaAdiantamento" & HID.Value) = Dr("ContaClientes")
    '        Next

    '        '--------------------------------------------------------------

    '        Cliente = DdlEmpresaCliente.SelectedValue
    '        campo = Cliente.Split("-")

    '        Sql = " INSERT INTO AdiantamentosXBaixas" & _
    '                    " (Empresa_ID " & _
    '                    ", EndEmpresa_ID" & _
    '                    ", Cliente_ID" & _
    '                    ", EndCliente_ID" & _
    '                    ", Adiantamento_Id" & _
    '                    ", Sequencia_Id" & _
    '                    ", RegistroPedido" & _
    '                    ", Titulo" & _
    '                    ", ValorOficial" & _
    '                    ", ValorMoeda" & _
    '                    ", VariacaoOficial" & _
    '                    ", VariacaoMoeda" & _
    '                    ", DataBaixa)" & _
    '                    " VALUES('" & campo(0) & "'" & _
    '                    "," & campo(1)

    '        Cliente = txtCodigoCliente.Value
    '        campo = Cliente.Split("-")
    '        Sql &= ", '" & campo(0) & "'"                                               'Cliente
    '        Sql &= ", " & campo(1)                                                      'Endereco Cliente

    '        Sql &= ", " & txtNumeroAdto.Text                                            'Numero do Adiantamentos
    '        Sql &= ", " & Sequencia                                                     'Sequencia

    '        Sql &= ", " & txtPedido.Text                                                'Numero do Pedido
    '        Sql &= ", " & txtRegistro.Text                                              'Titulo

    '        Sql &= ", " & Replace(Replace(txtValorDoDocumento.Text, ".", ""), ",", ".")
    '        Sql &= ", " & Replace(Replace(txtValorEmMoeda.Text, ".", ""), ",", ".")
    '        Sql &= ", 0"
    '        Sql &= ", 0"
    '        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "')"

    '        SqlArray.Add(Sql)
    '    End If
    '    Return True

    'End Function

    Private Sub Limpar(ByVal LimparConsulta As Boolean)
        Try
            Session.Remove("ssRegistros" & HID.Value)
            Session.Remove("ssObservacoes" & HID.Value)
            Session.Remove("ContaAdiantamento" & HID.Value)
            Session.Remove("ssGrupado" & HID.Value)
            Session.Remove("objPedidoSelecionadoCTAREC" & HID.Value)
            Session.Remove("objDestinoContabil" & HID.Value)
            Session.Remove("ControleCR" & HID.Value)

            If LimparConsulta AndAlso lblAgrupar.Text = "AP" Then LimparConsultaTitulos(False)

            If Not chkManterLancamento.Checked Then
                ddlMoeda.SelectedIndex = 0

                lblUsuarioIncl.Text = ""
                lblUsuarioAlt.Text = ""
                imgUsuarioIncl.Visible = False
                imgUsuarioAlt.Visible = False
                imgBloqueio.Visible = False
                imgExtratoPedido.Visible = False

                txtContratoBanco.Text = String.Empty

                txtCodigoFaturaDeFrete.Text = 0
                DivFaturaDeFrete.Style.Add("Display", "none")

                If Funcoes.VerificaPermissao("ContratoBancario", "GRAVAR") Then
                    rowContratoBanco.Visible = True
                Else
                    rowContratoBanco.Visible = False
                End If

                txtPedidoConsultaTitulos.Text = ""

                txtLiberarTitulo.Value = "N"
                txtLiberarPedido.Value = "N"
                txtUsuarioLiberarTitulo.Value = ""
                txtUsuarioLiberarPedido.Value = ""
                txtUsuarioLiberarTituloData.Value = ""
                txtUsuarioLiberarPedidoData.Value = ""

                txtCliente.Text = String.Empty
                txtCodigoCliente.Value = String.Empty
                DdlTributos.Items.Clear()

                DdlEmpresaPagadora.SelectedIndex = 0
                DdlBancoRecebedor.Items.Clear()
                DdlContaRecebedora.Items.Clear()

                ddlCarteiras.Enabled = True
                ddlCarteiras.SelectedIndex = 0
                DdlTiposDePagamentos.SelectedIndex = 0
                DdlProvisoes.SelectedIndex = 0
                ddlCarteiraDoTitulo.SelectedIndex = 0
                txtContratoFinanceiro.Text = ""
                txtPedidoFixacao.Value = ""

                ddlMoeda.SelectedValue = 1
                ddlIndexador.SelectedIndex = 3

                txtRegistro.Enabled = True

                txtPedido.Text = "0"

                ddlSelecionarHist.SelectedValue = String.Empty
                txtHistorico.Text = ""
                lblVencOriginal.Text = ""
                cmdPedido.Enabled = True
                cmdAdiantamento.Enabled = True

                lblAgrupar.Text = ""
                txtMestre.Text = ""
                txtObservacoes.Text = ""
                MensagemParcial = ""

                GridPedidos.DataSource = Nothing
                GridPedidos.DataBind()

                gridRazao.DataSource = Nothing
                gridRazao.DataBind()

                GridConsultaTitulos.Visible = True

                For Each objLinha As GridViewRow In GridConsultaTitulos.Rows
                    Dim chkTitulo As CheckBox = objLinha.Cells(0).FindControl("ChkGridTitulos")
                    If chkTitulo.Checked Then chkTitulo.Checked = False
                Next

                DdlUnidadeDeNegocioEmpresaCliente.Enabled = True
                DdlEmpresaCliente.Enabled = True
                ddlCarteiras.Enabled = True
                DdlProvisoes.Enabled = True
                txtHistorico.Enabled = True

                txtMovimento.Enabled = True
                ShowCalendar(Me.Page, txtMovimento)
                txtProrrogacao.Enabled = True
                ShowCalendar(Me.Page, txtProrrogacao)

                ddlMoeda.Enabled = True
                ddlIndexador.Enabled = True
                lblCotacao.Text = String.Empty
                lblDescCotacao.Text = String.Empty
                txtValorDoDocumento.Enabled = True
                txtDescontos.Enabled = True
                txtDeducoes.Enabled = True
                txtJuros.Enabled = True
                txtAcrescimos.Enabled = True
                txtValorCobrado.Enabled = True

                txtValorEmMoeda.Enabled = True
                txtDescontosMoeda.Enabled = True
                txtDeducoesMoeda.Enabled = True
                txtJurosMoeda.Enabled = True
                txtAcrescimosMoeda.Enabled = True
                txtValorCobradoMoeda.Enabled = True

                lnkNovo.Parent.Visible = False
                lnkExcluir.Parent.Visible = False

                If Not DdlEmpresaCliente.SelectedIndex > 0 Then VerificaUnidade()

                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaPedidos.SetarHID(HID.Value)
                ucConsultaAdiantamentos.SetarHID(HID.Value)
                ucDestinoContabil.SetarHID(HID.Value)
            End If

            txtNumeroAdto.Text = "0"
            HDSaldoAdiantamento.Value = 0

            Session("ssRegistros" & HID.Value) = ""
            Session("ssObservacoes" & HID.Value) = ""
            Session("ContaAdiantamento" & HID.Value) = ""
            Session("ssGrupado" & HID.Value) = ""

            txtProrrogacao.Text = ""
            hdnProrrogacaoOriginal.Value = String.Empty
            txtRegistro.Text = ""

            txtDescontos.Text = ""
            txtDeducoes.Text = ""
            txtJuros.Text = ""
            txtAcrescimos.Text = ""
            txtValorCobrado.Text = ""

            txtDescontosMoeda.Text = ""
            txtDeducoesMoeda.Text = ""
            txtJurosMoeda.Text = ""
            txtAcrescimosMoeda.Text = ""
            txtValorCobradoMoeda.Text = ""

            lblDeducoes.Text = "Deduções:"
            lblDescontos.Text = "Descontos:"
            lblJuros.Text = "Juros:"
            lblAcrescimos.Text = "Acréscimos:"

            txtValorDoDocumento.Text = ""
            txtValorDocumento.Value = 0
            txtValorMoeda.Value = 0
            txtValorEmMoeda.Text = ""
            txtValorLiquido.Value = 0
            txtValorLiquidoMoeda.Value = 0

            ImgCalcular.Enabled = True

            txtValorDoDocumento.ForeColor = Drawing.Color.Black
            txtValorEmMoeda.ForeColor = Drawing.Color.Black
            txtOficial.ForeColor = Drawing.Color.Black
            txtMoeda.ForeColor = Drawing.Color.Black

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkAllTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If GridConsultaTitulos.Rows.Count > 0 Then
                Dim chkAllTitulos As CheckBox = CType(sender, CheckBox)
                Dim passed As Boolean = False
                For Each rowgrid As GridViewRow In GridConsultaTitulos.Rows
                    Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("ChkGridTitulos"), CheckBox)

                    If Not chkTitulo.Enabled Then
                        chkAllTitulos.Checked = False
                        Exit Sub
                    End If

                    If chkAllTitulos.Checked Then
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

    Private Function verificaSelecionados() As Boolean
        Dim passed As Boolean = False
        For Each row As GridViewRow In GridConsultaTitulos.Rows
            Dim chkTitulo As CheckBox = CType(row.FindControl("ChkGridTitulos"), CheckBox)
            If chkTitulo.Checked Then
                passed = True
                Exit For
            End If
        Next
        Return passed
    End Function

    Protected Sub GridConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Registro = GridConsultaTitulos.SelectedRow.Cells(2).Text()
            Limpar(False)
            txtRegistro.Text = Registro
            ConsultaContasAReceber()
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Sub ConsultaCLientes(ByVal Cli As String, ByVal EndCli As Integer)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        Sql = "SELECT Cliente_Id as Codigo, Endereco_Id, Nome, Cidade, Estado" & vbCrLf & _
              " FROM Clientes Where Situacao IN (1,50) And Cliente_Id  = '" & Cli & "' And Endereco_Id = " & EndCli

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id"))
            txtCliente.Text = Descricao
            txtCodigoCliente.Value = Codigo
        Next

    End Sub

    Protected Sub cmdAlterar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        GravaTitulo()
    End Sub

    Sub ConsultaContasAReceber()
        Dim Pedido As Integer = 0
        Dim SaldoParcelas As Decimal
        Dim conta As String

        TemRegistro = ""
        Try
            If txtRegistro.Text <> "" Then

                Registro = txtRegistro.Text

                txtRegistro.Text = Registro

                Sql = " SELECT CR.Registro_Id, CR.Sequencia_Id, CR.Provisao, CR.Carteira, CR.Tributo, CR.Indexador, " & vbCrLf & _
                      "        CR.Moeda, CR.TipoPagto, CR.Situacao, CR.Lote, CR.Movimento, CR.Vencimento, CR.Prorrogacao, " & vbCrLf & _
                      "        CR.DataMoeda, CR.Baixa, CR.UnidadeDeNegocio, CR.Empresa, CR.EndEmpresa, CR.Cliente, " & vbCrLf & _
                      "        CR.EndCliente, CR.BancoCliente, CR.AgenciaCliente, CR.DigitoAgenciaCliente, CR.ContaCliente, " & vbCrLf & _
                      "        CR.DigitoContaCliente, Isnull(CR.TipoContaCliente,'C') AS TipoContaCliente, CR.ContaContabilCliente, " & vbCrLf & _
                      "        CR.EmpresaPagadora, CR.EndEmpresaPagadora, CR.BancoPagador, isnull(CR.AgenciaPagadora,'') AS AgenciaPagadora, CR.DigitoAgenciaPagadora, " & vbCrLf & _
                      "        CR.ContaPagadora, CR.DigitoContaPagadora, CR.ContaContabilPagadora, CR.Cheque, isnull(CR.Slips,'N') AS Slips, " & vbCrLf & _
                      "        CR.Recibo, CR.Aviso, CR.ReciboDeposito, isnull(CR.EmpresaPedido,'') AS EmpresaPedido, isnull(CR.EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(CR.Pedido, 0) AS Pedido, " & vbCrLf & _
                      "        isnull(CR.PedidoFixacao,0) AS PedidoFixacao, isnull(CR.Procuracao,0) AS Procuracao, CR.ValorDoDocumento, CR.Descontos, CR.Deducoes, " & vbCrLf & _
                      "        CR.Juros, CR.Acrescimos, CR.ValorLiquido, ISNULL(CR.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, " & vbCrLf & _
                      "        ISNULL(CR.MoedaDescontos, 0) AS MoedaDescontos, ISNULL(CR.MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(CR.MoedaJuros, 0) AS MoedaJuros, " & vbCrLf & _
                      "        ISNULL(CR.MoedaAcrescimos, 0) AS MoedaAcrescimos, ISNULL(CR.MoedaValorLiquido, 0) AS MoedaValorLiquido, CR.Historico, " & vbCrLf & _
                      "        CR.CodigoDeBarras, CR.CodigoDigitado, CR.Destinatario, CR.EndDestinatario, CR.NomeDoDestinatario, CR.Destinacao, " & vbCrLf & _
                      "        CR.Solicitacao, CR.UsuarioInclusao, CR.UsuarioInclusaoData, CR.UsuarioAlteracao, CR.UsuarioAlteracaoData, " & vbCrLf & _
                      "        CR.UsuarioCancelamento, CR.UsuarioCancelamentoData, isnull(CR.UsuarioLiberacao,'') AS UsuarioLiberacao, CR.UsuarioLiberacaoData, " & vbCrLf & _
                      "        CR.UsuarioBaixa, CR.UsuarioBaixaData, isnull(CR.Grupado,'N') AS Grupado, isnull(CR.RegistroMestre, 0) as RegistroMestre, CR.Observacoes, " & vbCrLf & _
                      "        CR.SituacaoBancaria, ISNULL(CR.NumeroDoCheque,0) AS NumeroDoCheque," & vbCrLf & _
                      "        isnull(CR.UsuarioLiberacaoBloqueio, '') AS UsuarioLiberacaoBloqueio, CR.UsuarioLiberacaoBloqueioDate, isnull(CR.UsuarioLiberacaoPedido, '') AS UsuarioLiberacaoPedido, " & vbCrLf & _
                      "        CR.UsuarioLiberacaoPedidoDate, isnull(CR.UsuarioLiberacaoCheque, '') AS UsuarioLiberacaoCheque, CR.UsuarioLiberacaoChequeDate, " & vbCrLf & _
                      "        isnull(CR.CarteiraAdto,'') AS CarteiraAdto, isnull(CR.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(CR.ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, " & vbCrLf & _
                      "        ISNULL(NF.TipoDeDocumento,0) AS TipoDeDocumento, isnull(Ped.FinanceiroAberto,1) AS FinanceiroAberto,  ISNULL(CR.ContratoBancario, '') AS ContratoBanco, " & vbCrLf & _
                      "        ISNULL(NF.Empresa_id, '') AS EmpresaNota, ISNULL(NF.EndEmpresa_Id,0) AS EndEmpresaNota, ISNULL(NF.Cliente_Id, '') AS ClienteNota, " & vbCrLf & _
                      "        ISNULL(NF.EndCliente_Id, 0) AS EndClienteNota, ISNULL(NF.EntradaSaida_Id, '') AS ESNota, ISNULL(NF.Serie_Id, 0) AS SerieNota, ISNULL(NF.Nota_Id, 0) AS Nota,  " & vbCrLf & _
                      "        ISNULL(FFxT.Fatura_Id, 0) AS FaturaDeFrete, " & vbCrLf & _
                      "        isnull(isnull(A.Adiantamento_id,Axb.Adiantamento_id),0) as Adiantamento," & vbCrLf & _
                      "        isnull(a.Vencimento,cr.Prorrogacao) as VencimentoAdto," & vbCrLf & _
                      "        isnull(a.Taxa,0) as TaxaAdto" & vbCrLf & _
                      "   FROM ContasAReceber CR " & vbCrLf & _
                      "   LEFT JOIN NotaFiscalXTitulo NFxT " & vbCrLf & _
                      "		ON CR.Registro_Id = NFxT.Titulo_Id " & vbCrLf & _
                      "   LEFT JOIN NotasFiscais NF " & vbCrLf & _
                      "		ON NFxT.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                      "	   AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                      "    AND NFxT.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                      "    AND NFxT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                      "    AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                      "    AND NFxT.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                      "    AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                       "  LEFT JOIN FaturaDeFreteXTitulo FFxT " & vbCrLf & _
                      "     ON CR.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
                      "	  LEFT JOIN Pedidos Ped " & vbCrLf & _
                      "     ON Ped.Empresa_Id    = CR.EmpresaPedido " & vbCrLf & _
                      "    AND Ped.EndEmpresa_Id = CR.EndEmpresaPedido " & vbCrLf & _
                      "    AND Ped.Pedido_Id     = CR.Pedido " & vbCrLf & _
                      "   Left Join Adiantamentos A" & vbCrLf & _
                      "     on A.Titulo = Cr.Registro_Id" & vbCrLf & _
                      "   Left Join AdiantamentosXBaixas AxB" & vbCrLf & _
                      "     on AxB.titulo = cr.Registro_Id" & vbCrLf & _
                      "  WHERE CR.Registro_Id    = " & Registro

                Dim dsContasAReceber As New DataSet
                dsContasAReceber = Banco.ConsultaDataSet(Sql, "ContasAReceber")

                If dsContasAReceber Is Nothing OrElse dsContasAReceber.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Registro não encontrado")
                    Exit Sub
                End If

                For Each Dr As DataRow In dsContasAReceber.Tables(0).Rows

                    If Dr("Pedido") > 0 AndAlso Dr("FinanceiroAberto") = "False" Then
                        MsgBox(Me.Page, "Título com Financeiro Fechado no Pedido não pode ser alterado")
                        lnkExcluir.Parent.Visible = False
                        lnkNovo.Parent.Visible = False
                    End If

                    Session("ControleCR" & HID.Value) = Dr("Registro_Id").ToString + ";" + Dr("ValorDoDocumento").ToString + ";" + Dr("MoedaValorDoDocumento").ToString + ";" + Dr("Moeda").ToString

                    DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = DdlUnidadeDeNegocioEmpresaCliente.Items.IndexOf(DdlUnidadeDeNegocioEmpresaCliente.Items.FindByValue(Dr("UnidadeDeNegocio")))
                    CargaEmpresaCliente()
                    DdlEmpresaCliente.SelectedIndex = DdlEmpresaCliente.Items.IndexOf(DdlEmpresaCliente.Items.FindByValue(Dr("Empresa") & "-" & CStr(Dr("EndEmpresa"))))
                    ConsultaCLientes(Dr("Cliente"), Dr("EndCliente"))

                    DdlEmpresaPagadora.SelectedIndex = DdlEmpresaPagadora.Items.IndexOf(DdlEmpresaPagadora.Items.FindByValue(Dr("EmpresaPagadora") & "-" & CStr(Dr("EndEmpresaPagadora"))))
                    DdlTiposDePagamentos.SelectedIndex = DdlTiposDePagamentos.Items.IndexOf(DdlTiposDePagamentos.Items.FindByValue(Dr("TipoPagto")))

                    If DdlEmpresaPagadora.Text <> "" Then
                        CargaBancoRecebedor()
                    End If

                    If Dr("BancoPagador") > 0 AndAlso Dr("AgenciaPagadora").ToString.Length > 0 Then
                        DdlBancoRecebedor.SelectedIndex = DdlBancoRecebedor.Items.IndexOf(DdlBancoRecebedor.Items.FindByValue(Dr("BancoPagador")))
                        ContaRecebedora()
                        conta = Dr("AgenciaPagadora") & "-" & Dr("DigitoAgenciaPagadora") & "-" & Dr("ContaPagadora") & "-" & Dr("DigitoContaPagadora") & "-" & Dr("ContaContabilPagadora")
                        DdlContaRecebedora.SelectedIndex = DdlContaRecebedora.Items.IndexOf(DdlContaRecebedora.Items.FindByValue(conta))
                    End If

                    ddlMoeda.SelectedValue = Dr("Moeda")
                    ddlIndexador.SelectedValue = Dr("Indexador") '(21-07-12 inclusao indexador na tela)

                    txtContratoFinanceiro.Text = Dr("ContratoDeFinanciamento")
                    ddlCarteiraDoTitulo.SelectedIndex = ddlCarteiraDoTitulo.Items.IndexOf(ddlCarteiraDoTitulo.Items.FindByValue(Dr("CarteiraDoTitulo")))

                    DdlProvisoes.SelectedIndex = DdlProvisoes.Items.IndexOf(DdlProvisoes.Items.FindByValue(Dr("Provisao")))
                    ddlCarteiras.SelectedIndex = ddlCarteiras.Items.IndexOf(ddlCarteiras.Items.FindByValue(Dr("Carteira")))


                    Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

                    If objCarteira.isAdiantamento And objCarteira.BaixaAdiantamento Then
                        If ddlMoeda.SelectedValue = 1 Then
                            HDSaldoAdiantamento.Value = Dr("ValorDoDocumento")
                        Else
                            HDSaldoAdiantamento.Value = Dr("MoedaValorDoDocumento")
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaDeducao) Then
                        lblDeducoes.Text = objCarteira.ContaDeducao.Titulo
                    Else
                        lblDeducoes.Text = "Deduções:"
                    End If
                    If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaDesconto) Then
                        lblDescontos.Text = objCarteira.ContaDesconto.Titulo
                    Else
                        lblDescontos.Text = "Descontos:"
                    End If
                    If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaJuro) Then
                        lblJuros.Text = objCarteira.ContaJuro.Titulo
                    Else
                        lblJuros.Text = "Juros:"
                    End If
                    If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaAcrescimo) Then
                        lblAcrescimos.Text = objCarteira.ContaAcrescimo.Titulo
                    Else
                        lblAcrescimos.Text = "Acréscimos:"
                    End If


                    If Not objCarteira.isAdiantamento Then
                        CargaTributos()

                        If Dr("Tributo").ToString.Length > 0 Then
                            DdlTributos.SelectedIndex = DdlTributos.Items.IndexOf(DdlTributos.Items.FindByValue(Dr("Tributo")))
                        End If
                    End If

                    '***************************************************************************************************************
                    '*****************************************  ADIANTAMENTOS ******************************************************
                    '***************************************************************************************************************
                    DdlTributos.Parent.Visible = Not objCarteira.isAdiantamento
                    txtNumeroAdto.Parent.Visible = objCarteira.isAdiantamento
                    txtVencimentoAdto.Parent.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

                    txtVencimentoAdto.Enabled = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)
                    txtTaxaAdto.Parent.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

                    If objCarteira.isAdiantamento Then txtNumeroAdto.Text = Dr("Adiantamento")
                    If objCarteira.isAdiantamento AndAlso Not objCarteira.BaixaAdiantamento Then
                        txtVencimentoAdto.Text = Dr("VencimentoAdto")
                        txtTaxaAdto.Text = Dr("TaxaAdto")
                    End If

                    '***************************************************************************************************************
                    '***************************************************************************************************************


                    If Dr("UsuarioAlteracao").ToString.Length > 0 Then
                        lblUsuarioIncl.Text = Dr("UsuarioInclusao") & " - " & CDate(Dr("UsuarioInclusaoData")).ToString("dd/MM/yyyy")
                        lblUsuarioAlt.Text = Dr("UsuarioAlteracao") & " - " & CDate(Dr("UsuarioAlteracaoData")).ToString("dd/MM/yyyy")
                    Else
                        lblUsuarioIncl.Text = Dr("UsuarioInclusao") & " - " & CDate(Dr("UsuarioInclusaoData")).ToString("dd/MM/yyyy")
                    End If

                    imgUsuarioIncl.Visible = True
                    imgUsuarioAlt.Visible = True

                    txtRegistro.Text = Dr("Registro_Id")
                    txtMestre.Text = "Mestre : " & Dr("RegistroMestre")

                    txtHistorico.Text = Dr("Historico")

                    txtMovimento.Text = CDate(Dr("Movimento")).ToString("dd/MM/yyyy")
                    hdnMovimentoOriginal.Value = CDate(Dr("Movimento")).ToString("dd/MM/yyyy")
                    txtProrrogacao.Text = CDate(Dr("Prorrogacao")).ToString("dd/MM/yyyy") '************************ Vencimento/Prorrogacao
                    hdnProrrogacaoOriginal.Value = CDate(Dr("Prorrogacao")).ToString("dd/MM/yyyy") '************************ Vencimento/Prorrogacao original

                    lblVencOriginal.Text = CDate(Dr("Vencimento")).ToString("dd/MM/yyyy") '************************ Vencimento Original

                    If IsDBNull(Dr("MoedaValorDoDocumento")) Then
                        Dr("MoedaValorDoDocumento") = 0
                    End If

                    If Not IsDBNull(Dr("Observacoes")) Then
                        txtObservacoes.Text = Dr("Observacoes")
                    Else
                        txtObservacoes.Text = ""
                    End If

                    If IsDBNull(Dr("Pedido")) Then
                        Dr("Pedido") = 0
                    End If

                    txtPedido.Text = Dr("Pedido")
                    Pedido = Dr("Pedido")
                    txtPedidoFixacao.Value = Dr("PedidoFixacao")

                    TemRegistro = "S"

                    Sql = "SELECT ISNULL(IndiceFixado,0) AS IndiceFixado, DataPedido " & vbCrLf & _
                          "  FROM Pedidos " & vbCrLf & _
                          " WHERE Pedido_id = " & Pedido & vbCrLf
                    Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql, "Pedidos")

                    If Dr("Moeda") = 1 Then
                        txtValorDoDocumento.ForeColor = Drawing.Color.Red
                        txtOficial.ForeColor = Drawing.Color.Blue
                    End If

                    If Dr("Moeda") = 3 Then
                        txtValorEmMoeda.ForeColor = Drawing.Color.Red
                        txtMoeda.ForeColor = Drawing.Color.Blue
                        If Dr("Provisao") = 2 Then
                            If Pedido > 0 Then
                                If dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                                    Dr("ValorLiquido") = Math.Round(Dr("MoedaValorDoDocumento") * dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                                Else
                                    Dr("ValorLiquido") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorDoDocumento"), Dr("Indexador"), dsPedido.Tables(0).Rows(0).Item("DataPedido"))
                                End If

                                If (Dr("ValorDoDocumento") - Dr("ValorLiquido")) < 0 Then
                                    Dr("Acrescimos") = (Dr("ValorDoDocumento") - Dr("ValorLiquido")) * -1
                                Else
                                    Dr("Deducoes") = (Dr("ValorDoDocumento") - Dr("ValorLiquido"))
                                End If
                            Else
                                Dr("ValorDoDocumento") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorDoDocumento"), Dr("Indexador"), Dr("Vencimento"))
                                Dr("ValorLiquido") = Dr("ValorDoDocumento")
                            End If
                        End If
                    End If

                    'Valores em Reais
                    txtValorDocumento.Value = CDec(Dr("ValorDoDocumento")).ToString("N2")
                    txtValorDoDocumento.Text = CDec(Dr("ValorDoDocumento")).ToString("N2")
                    txtDescontos.Text = CDec(Dr("Descontos")).ToString("N2")
                    txtDeducoes.Text = CDec(Dr("Deducoes")).ToString("N2")
                    txtJuros.Text = CDec(Dr("Juros")).ToString("N2")
                    txtAcrescimos.Text = CDec(Dr("Acrescimos")).ToString("N2")
                    txtValorLiquido.Value = CDec(Dr("ValorLiquido")).ToString("N2")

                    If Dr("Provisao") = 1 OrElse (Dr("Provisao") = 2 And Pedido > 0) Then
                        txtValorCobrado.Text = CDec(Dr("ValorLiquido")).ToString("N2")
                    Else
                        txtValorCobrado.Text = CDec(Dr("ValorDoDocumento") + Dr("Juros") + Dr("Acrescimos") - Dr("Descontos") - Dr("Deducoes")).ToString("N2")
                    End If

                    'Valores em Dólar
                    txtValorMoeda.Value = CDec(Dr("MoedaValorDoDocumento")).ToString("N2")
                    txtValorEmMoeda.Text = CDec(Dr("MoedaValorDoDocumento")).ToString("N2")
                    txtDescontosMoeda.Text = CDec(Dr("MoedaDescontos")).ToString("N2")
                    txtDeducoesMoeda.Text = CDec(Dr("MoedaDeducoes")).ToString("N2")
                    txtJurosMoeda.Text = CDec(Dr("MoedaJuros")).ToString("N2")
                    txtAcrescimosMoeda.Text = CDec(Dr("MoedaAcrescimos")).ToString("N2")
                    txtValorLiquidoMoeda.Value = CDec(Dr("MoedaValorLiquido")).ToString("N2")

                    If Dr("Provisao") = 1 OrElse (Dr("Provisao") = 2 And Pedido > 0) Then
                        txtValorCobradoMoeda.Text = CDec(Dr("MoedaValorLiquido")).ToString("N2")
                    Else
                        txtValorCobradoMoeda.Text = CDec(Dr("MoedaValorDoDocumento") + Dr("MoedaJuros") + Dr("MoedaAcrescimos") - Dr("MoedaDescontos") - Dr("MoedaDeducoes")).ToString("N2")
                    End If

                    If Dr("Moeda") = 1 AndAlso CDec(txtValorEmMoeda.Text) = 0 Then ValidaValores(True)

                    lnkNovo.Parent.Visible = True
                    lnkExcluir.Parent.Visible = True
                    lnkRelatorio.Parent.Visible = True

                    '******* Se Existir Nota do título e a empresa usuário for Insol, pode conter dados no campo"
                    txtContratoBanco.Text = Dr("ContratoBanco")

                    If CInt(Dr("FaturaDeFrete")) > 0 Then
                        txtCodigoFaturaDeFrete.Text = CInt(Dr("FaturaDeFrete"))
                        DivFaturaDeFrete.Style.Clear()
                        ddlCarteiras.Enabled = False
                    End If


                    If Dr("TipoDeDocumento") = 4 AndAlso Dr("Provisao") = 1 Then
                        lnkNovo.Parent.Visible = False
                        lnkExcluir.Parent.Visible = False
                        txtMovimento.Enabled = False
                        HideCalendar(Me.Page, txtMovimento)
                        txtProrrogacao.Enabled = False
                        HideCalendar(Me.Page, txtProrrogacao)
                        ImgCalcular.Enabled = False
                    ElseIf Pedido > 0 Then
                        cmdPedido.Enabled = False
                        imgExtratoPedido.Visible = True
                        ddlMoeda.Enabled = False
                        ddlIndexador.Enabled = False
                        lnkExcluir.Parent.Visible = False

                        If Dr("Grupado") = "S" Then
                            lnkNovo.Parent.Visible = False
                            DdlProvisoes.Enabled = False
                            ImgCalcular.Enabled = False
                            txtMovimento.Enabled = False
                            HideCalendar(Me.Page, txtMovimento)
                            txtProrrogacao.Enabled = False
                            HideCalendar(Me.Page, txtProrrogacao)
                        ElseIf Dr("Grupado") = "M" Then
                            DdlProvisoes.Enabled = False
                            lnkNovo.Parent.Visible = False
                            ImgCalcular.Enabled = False
                            imgBloqueio.Visible = True
                            txtMovimento.Enabled = False
                            HideCalendar(Me.Page, txtMovimento)
                            txtProrrogacao.Enabled = False
                            HideCalendar(Me.Page, txtProrrogacao)
                            If Dr("Provisao") = 2 Then lnkExcluir.Parent.Visible = True
                        ElseIf Dr("Provisao") = 1 Then
                            DdlProvisoes.Enabled = False
                            lnkNovo.Parent.Visible = False
                            ImgCalcular.Enabled = False
                            imgBloqueio.Visible = True
                            txtMovimento.Enabled = False
                            HideCalendar(Me.Page, txtMovimento)
                            txtProrrogacao.Enabled = False
                            HideCalendar(Me.Page, txtProrrogacao)
                        Else
                            DdlProvisoes.Enabled = True
                            imgBloqueio.Visible = False
                        End If
                    Else
                        cmdPedido.Enabled = True

                        If Dr("Grupado") = "S" Then
                            lnkNovo.Parent.Visible = False
                            ImgCalcular.Enabled = False
                            lnkExcluir.Parent.Visible = False
                            DdlProvisoes.Enabled = False
                            txtMovimento.Enabled = False
                            HideCalendar(Me.Page, txtMovimento)
                            txtProrrogacao.Enabled = False
                            HideCalendar(Me.Page, txtProrrogacao)
                        ElseIf Dr("Grupado") = "M" Then
                            DdlProvisoes.Enabled = False
                            lnkNovo.Parent.Visible = False
                            ImgCalcular.Enabled = False
                            imgBloqueio.Visible = True
                            txtMovimento.Enabled = False
                            HideCalendar(Me.Page, txtMovimento)
                            txtProrrogacao.Enabled = False
                            HideCalendar(Me.Page, txtProrrogacao)
                            If Dr("Provisao") = 2 Then lnkExcluir.Parent.Visible = True
                        ElseIf Dr("Provisao") = 1 Then
                            DdlProvisoes.Enabled = False
                            lnkNovo.Parent.Visible = False
                            ImgCalcular.Enabled = False
                            imgBloqueio.Visible = True
                            txtMovimento.Enabled = False
                            HideCalendar(Me.Page, txtMovimento)
                            txtProrrogacao.Enabled = False
                            HideCalendar(Me.Page, txtProrrogacao)
                        Else
                            DdlProvisoes.Enabled = True
                            imgBloqueio.Visible = False
                        End If
                    End If

                    If Dr("Grupado") = "M" Then lblAgrupar.Text = "AP"

                    If lblAgrupar.Text = "AP" Then
                        cmdPedido.Enabled = False
                        Dim Filhos As New ArrayList
                        Dim Mensagem As String = "Agrupamento dos Registros"

                        Sql = " SELECT Registro_id " & vbCrLf & _
                              "   FROM ContasAreceber " & vbCrLf & _
                              "  WHERE RegistroMestre = " & txtRegistro.Text

                        Dim dsFilhos As New DataSet
                        dsFilhos = Banco.ConsultaDataSet(Sql, "RegistrosFilhos")

                        If Not dsFilhos Is Nothing AndAlso dsFilhos.Tables(0).Rows.Count > 0 Then
                            For Each drFilho As DataRow In dsFilhos.Tables(0).Rows
                                Filhos.Add(drFilho("Registro_id"))
                                Mensagem &= " - " & drFilho("Registro_id")
                            Next

                            Session("ssRegistros" & HID.Value) = Filhos
                            Session("ssObservacoes" & HID.Value) = Mensagem
                        End If
                    End If

                    If IsDBNull(Dr("UsuarioLiberacaoBloqueio")) OrElse Dr("UsuarioLiberacaoBloqueio") = "" Then
                        txtUsuarioLiberarTitulo.Value = ""
                        txtUsuarioLiberarTituloData.Value = CDate(Today).ToString("yyyy/MM/dd")
                    Else
                        txtUsuarioLiberarTitulo.Value = Dr("UsuarioLiberacaoBloqueio")
                        txtUsuarioLiberarTituloData.Value = Dr("UsuarioLiberacaoBloqueioDate")
                    End If

                    If IsDBNull(Dr("UsuarioLiberacaoPedido")) OrElse Dr("UsuarioLiberacaoPedido") = "" Then
                        txtUsuarioLiberarPedido.Value = ""
                        txtUsuarioLiberarPedidoData.Value = CDate(Today).ToString("yyyy/MM/dd")
                    Else
                        txtUsuarioLiberarPedido.Value = Dr("UsuarioLiberacaoPedido")
                        txtUsuarioLiberarPedidoData.Value = Dr("UsuarioLiberacaoPedidoDate")
                    End If
                Next

                'If TemRegistro = "S" Then
                If Pedido > 0 Then
                    Codigo = txtCodigoCliente.Value
                    Descricao = txtCliente.Text

                    Sql = "  SELECT Registro_Id AS Registro, Vencimento, Baixa, Historico, ValorDoDocumento, Descontos, " & vbCrLf & _
                          "         Deducoes, Juros, Acrescimos, ValorLiquido, 0.0 as Saldo, Provisao " & vbCrLf & _
                          "    FROM ContasAReceber " & vbCrLf & _
                          "   WHERE Pedido = '" & Pedido & "' " & vbCrLf & _
                          "     AND Situacao = 1 " & vbCrLf & _
                          "   ORDER BY Vencimento" & vbCrLf

                    Dim ds As New DataSet
                    Dim dra As DataRow
                    ds = Banco.ConsultaDataSet(Sql, "Pedidos")

                    For Each dra In ds.Tables(0).Rows
                        If dra("Provisao") = 2 Then
                            dra("ValorLiquido") = 0
                        End If

                        SaldoParcelas += dra("ValorDoDocumento") - (dra("ValorLiquido") + dra("Descontos") + dra("Deducoes") - dra("Juros") - dra("Acrescimos"))
                        dra("Saldo") = SaldoParcelas
                    Next

                    GridPedidos.DataSource = ds
                    GridPedidos.DataBind()

                    DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
                    DdlEmpresaCliente.Enabled = False
                    cmdPedido.Enabled = False
                    TabContainer1.ActiveTabIndex = 0
                End If


                Sql = "SELECT R.Empresa_Id + '-' + cast(R.EndEmpresa_Id as varchar) as Empresa, R.Conta_Id, R.Cliente_Id, R.EndCliente_Id, PlanoDeContas.Titulo, R.Movimento_Id, R.Lote_Id, " & vbCrLf & _
                      "       ISNULL(R.Produto, '') AS Produto, R.Custo, R.Historico, R.DebitoOficial, R.CreditoOficial " & vbCrLf & _
                      "  FROM Razao R" & vbCrLf & _
                      " INNER JOIN PlanoDeContas " & vbCrLf & _
                      "    ON R.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf & _
                      " WHERE R.Titulo = " & Registro & vbCrLf

                Dim dsRazao As New DataSet
                dsRazao = Banco.ConsultaDataSet(Sql, "Razao")

                If Not dsRazao Is Nothing AndAlso dsRazao.Tables(0).Rows.Count > 0 Then
                    Dim dtRazao As New DataTable("razao")
                    Dim saldo As Double = 0

                    dtRazao.Columns.Add("Empresa", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Conta", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Cliente", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Movimento", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Titulo", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Lote", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Produto", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Custo", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Historico", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Debito", Type.GetType("System.Double"))
                    dtRazao.Columns.Add("Credito", Type.GetType("System.Double"))
                    dtRazao.Columns.Add("Saldo", Type.GetType("System.Double"))

                    For Each row As DataRow In dsRazao.Tables(0).Rows
                        Dim drRazao As DataRow = dtRazao.NewRow()

                        drRazao("Empresa") = row("Empresa")
                        drRazao("Conta") = row("Conta_Id")
                        drRazao("Cliente") = row("Cliente_Id") & "-" & row("EndCliente_Id")
                        drRazao("Movimento") = row("Movimento_Id")
                        drRazao("Titulo") = row("Titulo")
                        drRazao("Lote") = row("Lote_Id")
                        drRazao("Produto") = row("Produto")
                        drRazao("Custo") = row("Custo")
                        drRazao("Historico") = row("Historico")
                        drRazao("Debito") = row("DebitoOficial")
                        drRazao("Credito") = row("CreditoOficial")
                        saldo = Math.Round(saldo + (row("DebitoOficial") - row("CreditoOficial")), 2)
                        drRazao("Saldo") = saldo
                        dtRazao.Rows.Add(drRazao)
                    Next
                    gridRazao.DataSource = dtRazao
                    gridRazao.DataBind()
                End If

                If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                    Sql = "SELECT Registro_Id, Cliente, EndCliente, Pedido, Carteira " & vbCrLf & _
                          "  FROM TitulosXDesdobrarFornecedor " & vbCrLf & _
                          " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

                    Dim dsDesdobrarFornecedor As New DataSet
                    dsDesdobrarFornecedor = Banco.ConsultaDataSet(Sql, "DesdobrarFornecedor")

                    If Not dsDesdobrarFornecedor Is Nothing AndAlso dsDesdobrarFornecedor.Tables(0).Rows.Count > 0 Then
                        Session("objDestinoContabil" & HID.Value) = dsDesdobrarFornecedor.Tables(0).Rows(0).Item("Cliente") & "-" & _
                                                        dsDesdobrarFornecedor.Tables(0).Rows(0).Item("EndCliente") & "-" & _
                                                        dsDesdobrarFornecedor.Tables(0).Rows(0).Item("Carteira")
                    End If
                End If

                txtRegistro.Enabled = False

                If DdlProvisoes.SelectedValue = 3 Then lnkNovo.Parent.Visible = False 'Provisao Bloqueada P/ Alteracao.

                If Not dsContasAReceber Is Nothing AndAlso dsContasAReceber.Tables(0).Rows.Count > 0 AndAlso Not dsContasAReceber.Tables(0).Rows(0).Item("Situacao") = 1 Then
                    lnkNovo.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    DdlProvisoes.Enabled = False
                    lnkRelatorio.Parent.Visible = False
                    imgBloqueio.Visible = False
                    txtMovimento.Enabled = False
                    HideCalendar(Me.Page, txtMovimento)
                    txtProrrogacao.Enabled = False
                    HideCalendar(Me.Page, txtProrrogacao)
                    txtMestre.Text = "CANCELADO"
                End If

                TabContainer1.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Informe o Numero do Registro para consulta...")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(RTrim(ex.ToString)))
        End Try
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaConsultaTitulos.Items.Clear()

        Sql = "  SELECT Cli .Cliente_Id as Codigo, Cli .Endereco_Id, Cli .Reduzido, Cli .Nome, Cli .Cidade, Cli .Estado " & vbCrLf & _
              "    FROM GruposXEmpresas GxE " & vbCrLf & _
              "   INNER JOIN Clientes Cli " & vbCrLf & _
              "      ON GxE.Cliente_Id = Cli .Cliente_Id " & vbCrLf & _
              "     AND GxE.EndCliente_Id = Cli .Endereco_Id" & vbCrLf & _
              "   WHERE GxE.Empresa_Id = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresaConsultaTitulos.Items.Insert(0, "")
        DdlEmpresaConsultaTitulos.SelectedIndex = 0

    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objConsultaCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(obj)
                txtCodigoCliente.Value = itemCliente.Value
                txtCliente.Text = itemCliente.Text
                Session.Remove("objConsultaCliente" & HID.Value)
            ElseIf Not Session("objConsultaClienteTitulos" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(obj)
                txtCodigoClienteConsTitulo.Value = itemCliente.Value
                txtClienteConsTitulo.Text = itemCliente.Text
                Session.Remove("objConsultaClienteTitulos" & HID.Value)
            ElseIf Session("objContasAReceber" & HID.Value) IsNot Nothing Then
                Dim p As [Lib].Negocio.Pedido = CType(Session("objContasAReceber" & HID.Value), [Lib].Negocio.Pedido)
                txtPedidoConsultaTitulos.Text = p.Codigo
                Session.Remove("objContasAReceber" & HID.Value)
            ElseIf Not Session("objPedidoCTAREC" & HID.Value) Is Nothing Then
                Cliente = DdlEmpresaCliente.SelectedValue
                campo = Cliente.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                objPedido = CType(Session("objPedidoCTAREC" & HID.Value), [Lib].Negocio.Pedido)
                If Trim(txtValorEmMoeda.Text) = "" Then txtValorEmMoeda.Text = 0
                If objPedido.MomentoFinanceiro = 3 And Funcoes.VerificaPermissao("ContasAReceber", "LIBERARMOMENTOFINANCEIRO") = False And objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                    MsgBox(Me.Page, "Processo Não permitido. Pedido Lançado Com Vencimentos Determinados na Emissão da Nota Fiscal.")
                ElseIf objPedido.CodigoUnidadeNegocio <> DdlUnidadeDeNegocioEmpresaCliente.SelectedValue Then
                    MsgBox(Me.Page, "Unidade de Negócio da Empresa do Pedido é diferente da Unidade de Negócio da Empresa do Cliente.")
                ElseIf objPedido.CodigoEmpresa <> campo(0) Or objPedido.EnderecoEmpresa.ToString <> campo(1) Then
                    MsgBox(Me.Page, "Empresa do Pedido é diferente da Empresa do Cliente.")
                ElseIf objPedido.CodigoCliente <> strCliente(0) Or objPedido.EnderecoCliente.ToString <> strCliente(1) Then
                    MsgBox(Me.Page, "Cliente do Pedido é diferente do Cliente informado.")
                Else
                    txtPedido.Text = objPedido.Codigo
                    txtVencimentoAdto.Text = CDate(objPedido.DataPedido).ToString("dd/MM/yyyy")
                    ddlMoeda.SelectedValue = objPedido.CodigoMoeda
                    ddlIndexador.SelectedValue = objPedido.CodigoIndexador
                    ddlMoeda.Enabled = False
                    ddlIndexador.Enabled = False
                End If
                Session.Remove("objPedidoCTAREC" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdClientesTitulo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClientesTitulo.Click
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objConsultaCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objConsultaClienteTitulos" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAgrupar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAgrupar.Click
        Dim Mensagem As String = "Agrupamento dos Registros"
        lblAgrupar.Text = "AP"
        Dim TestaCliente As String = ""
        Dim ValidaCliente As String = "S"
        Dim Registros As New ArrayList
        txtRegistro.Text = String.Empty


        Sql = "SELECT UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, Indexador, " & vbCrLf & _
              "       Sum(ValorLiquido) as ValorLiquido, " & vbCrLf & _
              "       Sum(isnull(MoedaValorLiquido,0)) as MoedaValorLiquido" & vbCrLf & _
              "  FROM ContasAReceber " & vbCrLf & _
              " WHERE Registro_Id = 99999999 " & vbCrLf

        Dim i As Integer = 0
        While i < GridConsultaTitulos.Rows.Count
            If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                Mensagem &= " - " & GridConsultaTitulos.Rows(i).Cells(2).Text()
                Sql &= " or Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(2).Text()
                Registros.Add(GridConsultaTitulos.Rows(i).Cells(2).Text())
            End If
            i += 1
        End While

        Sql &= " Group By UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, Indexador "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "ContasAReceber").Tables(0).Rows
            If TestaCliente <> "" Then
                If TestaCliente <> Dr("Cliente") Then
                    ValidaCliente = "N"
                    Exit For
                End If
            End If

            DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = DdlUnidadeDeNegocioEmpresaCliente.Items.IndexOf(DdlUnidadeDeNegocioEmpresaCliente.Items.FindByValue(Dr("UnidadeDeNegocio")))
            CargaEmpresaCliente()
            DdlEmpresaCliente.SelectedIndex = DdlEmpresaCliente.Items.IndexOf(DdlEmpresaCliente.Items.FindByValue(Dr("Empresa") & "-" & CStr(Dr("EndEmpresa"))))
            ConsultaCLientes(Dr("Cliente"), Dr("EndCliente"))
            DdlProvisoes.SelectedIndex = 1

            Codigo = txtCodigoCliente.Value
            Descricao = txtCliente.Text

            txtMovimento.Text = CDate(Today).ToString("dd/MM/yyyy")
            hdnMovimentoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")
            txtProrrogacao.Text = CDate(Today).ToString("dd/MM/yyyy")
            hdnProrrogacaoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")

            ddlMoeda.SelectedValue = txtRealDolar.Value
            ddlIndexador.SelectedValue = HiddenIndexador.Value
            ddlMoeda.Enabled = False
            ddlIndexador.Enabled = False

            If ddlMoeda.SelectedValue = 1 Then
                If ddlIndexador.SelectedValue = 99 Then
                    Dr("MoedaValorLiquido") = 0
                Else
                    Dr("MoedaValorLiquido") = Funcoes.ConverteParaMoedaExtrangeira(Dr("ValorLiquido"), ddlIndexador.SelectedValue, CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd"))
                End If
            End If

            If ddlMoeda.SelectedValue = 3 Then
                If ddlIndexador.SelectedValue = 99 Then
                    Dr("ValorLiquido") = 0
                Else
                    Dr("ValorLiquido") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorLiquido"), ddlIndexador.SelectedValue, CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd"))
                End If
            End If

            txtValorDoDocumento.Text = CDec(Dr("ValorLiquido")).ToString("N2")
            txtValorCobrado.Text = CDec(Dr("ValorLiquido")).ToString("N2")

            txtValorEmMoeda.Text = CDec(Dr("MoedaValorLiquido")).ToString("N2")

            txtDescontos.Text = "0,00"
            txtDeducoes.Text = "0,00"
            txtJuros.Text = "0,00"
            txtAcrescimos.Text = "0,00"


            txtValorDoDocumento.Enabled = False
            txtValorEmMoeda.Enabled = False
            txtDescontos.Enabled = False
            txtDeducoes.Enabled = False
            txtJuros.Enabled = False
            txtAcrescimos.Enabled = False
            txtValorCobrado.Enabled = False

            TestaCliente = Dr("Cliente")
        Next

        If ValidaCliente = "S" Then
            Dim objCliente As New [Lib].Negocio.Cliente(txtCodigoCliente.Value.ToString.Split("-")(0), txtCodigoCliente.Value.ToString.Split("-")(1))
            txtHistorico.Text = "RECBTO. " & objCliente.Nome

            Session("ssRegistros" & HID.Value) = Registros
            Session("ssObservacoes" & HID.Value) = Mensagem
            TabPanel4.Visible = False
            TabContainer1.ActiveTabIndex = 0
            Session("ssGrupado" & HID.Value) = "S"
        Else
            Limpar(True)
            MsgBox(Me.Page, "Agrupamento só pode ser realizado se for para o mesmo cliente.")
        End If
    End Sub

    Private Sub LimparConsultaTitulos(ByVal limparTudo As Boolean)
        If limparTudo Then
            If Not DdlEmpresaConsultaTitulos.SelectedIndex > 0 Then Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
            txtCodigoClienteConsTitulo.Value = String.Empty
            txtClienteConsTitulo.Text = String.Empty
            txtPeriodoInicialConsultaTitulos.Text = String.Format("{0}/{1}/{2}", "01", Month(Now).ToString("00"), Year(Now))
            txtPeriodoFinalConsultaTitulos.Text = CDate(Now).ToString("dd/MM/yyyy")
        End If

        lblTotalRegistroAgrupado.Parent.Visible = False
        lblTotalRegistroAgrupado.Text = String.Empty

        lnkAgrupar.Parent.Visible = False
        lnkBaixar.Parent.Visible = False
        lnkEmiteReciboGeral.Parent.Visible = False
        lnkReprogramar.Parent.Visible = False
        pnlReprogramaVencimentos.Visible = False
        txtNovoVencimento.Text = String.Empty
        GridConsultaTitulos.DataBind()
    End Sub

    Protected Sub lnkLimparConsultaTitulos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            LimparConsultaTitulos(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBaixar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim Mensagem As String = "Baixa dos Registros"
        'Dim i As Integer = 0
        'Dim valor As String = 0
        'lblAgrupar.Text = "BT"

        'Dim Registros As New ArrayList

        'Sql = "SELECT Sum(ValorDoDocumento) as ValorDoDocumento, Sum(Descontos) as Descontos, Sum(Deducoes) as Deducoes, Sum(Juros) as Juros, Sum(Acrescimos) as Acrescimos, Sum(ValorLiquido) as ValorLiquido "
        'Sql &= " FROM ContasAReceber"
        'Sql &= " WHERE Registro_Id = 99999999 "

        'While i < GridConsultaTitulos.Rows.Count
        '    If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
        '        Mensagem &= " - " & GridConsultaTitulos.Rows(i).Cells(2).Text()
        '        Sql &= " or Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(2).Text()
        '        Registros.Add(GridConsultaTitulos.Rows(i).Cells(2).Text())
        '    End If
        '    i += 1
        'End While

        'For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "ContasAReceber").Tables(0).Rows
        '    Valor = Dr("ValorDoDocumento")
        'Next

        'If Valor > 0 Then
        '    For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "ContasAReceber").Tables(0).Rows
        '        DdlProvisoes.SelectedIndex = 1

        '        txtValorDoDocumento.Text = Dr("ValorDoDocumento")
        '        txtDescontos.Text = Dr("Descontos")
        '        txtDeducoes.Text = Dr("Deducoes")
        '        txtJuros.Text = Dr("Juros")
        '        txtAcrescimos.Text = Dr("Acrescimos")
        '        txtValorCobrado.Text = Dr("ValorLiquido")

        '        txtMovimento.Text = CDate(Today).ToString("dd/MM/yyyy")
        '        txtVencimento.Text = CDate(Today).ToString("dd/MM/yyyy")
        '    Next

        '    HttpContext.Current.Session("ssRegistros") = Registros
        '    HttpContext.Current.Session("ssObservacoes") = Mensagem


        '    DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
        '    DdlEmpresaCliente.Enabled = False
        '    DdlCliente.Enabled = False
        '    ddlCarteiras.Enabled = False
        '    DdlProvisoes.Enabled = False
        '    txtHistorico.Enabled = False

        '    TabContainer1.ActiveTabIndex = 0
        'Else
        '    Mensagem = "Nenhum Registro foi selecionado..."
        MsgBox(Me.Page, "Opção ainda não Liberada...")
        'End If
    End Sub

    Protected Sub DdlBancoRecebedor_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ContaRecebedora(True)
    End Sub


    Sub ContaRecebedora(Optional ByVal SomenteAtivas As Boolean = False)
        Dim Cliente As String
        Dim Campo() As String
        Dim Conta As String

        Cliente = DdlEmpresaPagadora.SelectedValue
        Campo = Cliente.Split("-")
        DdlContaRecebedora.Items.Clear()

        If Campo(0) <> "" Then
            Sql = " SELECT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes " & vbCrLf & _
                  "   FROM BancosXContas" & vbCrLf & _
                  "  WHERE Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id  = " & Campo(1) & vbCrLf & _
                  "    AND Banco_Id  = " & IIf(DdlBancoRecebedor.SelectedValue = "", "0", DdlBancoRecebedor.SelectedValue) & vbCrLf
            If SomenteAtivas Then
                Sql &= "   AND Ativo = 1 " & vbCrLf
            End If

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("ContaContabil")
                Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "  C/C " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes")
                DdlContaRecebedora.Items.Add(New ListItem(Descricao, Conta))
            Next
            DdlContaRecebedora.Items.Insert(0, "")
            DdlContaRecebedora.SelectedIndex = 0
        End If
    End Sub

    Protected Sub ImgCalcular_Click1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ValidaValores(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidaData(ByVal Data As String, ByVal Tipo As String, ByVal Empresa As String, ByVal EndEmpresa As String) As Boolean
        If Not IsDate(Data) Then
            MsgBox(Me.Page, " Data de " & Tipo & " inválida...")
            Return False
        ElseIf CDate(Data).DayOfWeek = 6 Then
            MsgBox(Me.Page, "Sábado - Data Inválida para " & Tipo & "...")
            Return False
        ElseIf CDate(Data).DayOfWeek = 0 Then
            MsgBox(Me.Page, "Domingo - Data Inválida para " & Tipo & "...")
            Return False
        End If

        Sql = "  SELECT Descricao" & vbCrLf & _
              "    FROM DatasNaoProgramaveis" & vbCrLf & _
              "   WHERE Empresa_Id = '99999999999999' " & vbCrLf & _
              "     AND EndEmpresa_ID = 0 " & vbCrLf & _
              "     AND Data_ID = '" & CDate(Data).ToString("yyyy/MM/dd") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            MsgBox(Me.Page, "Data de " & Tipo & " não programável, Feriado Nacional > " & Dr("Descricao"))
            Return False
        Next

        If Not String.IsNullOrWhiteSpace(Empresa) Then
            Sql = "  SELECT Descricao" & vbCrLf & _
                  "    FROM DatasNaoProgramaveis" & vbCrLf & _
                  "   WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                  "     AND EndEmpresa_ID = " & EndEmpresa & vbCrLf & _
                  "     AND Data_ID = '" & CDate(Data).ToString("yyyy/MM/dd") & "'" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                MsgBox(Me.Page, "Data de " & Tipo & " não programável, Feriado Municipal > " & Dr("Descricao"))
                Return False
            Next
        End If

        If Tipo.Equals("Movimento") Then
            If Not Funcoes.VerificaAcesso(Empresa, EndEmpresa, Data, "Financeiro") Then
                MsgBox(Me.Page, "Movimento já Fechado para esta data " & Data)
                Return False
            End If
        End If

        Return True
    End Function

    Sub Dolariza()
        If DdlProvisoes.Text <> "" Then
            If txtValorDoDocumento.Text = "" Then
                txtValorDoDocumento.Text = 0
            End If
            If txtValorEmMoeda.Text = "" Then
                txtValorEmMoeda.Text = 0
            End If

            If ddlMoeda.SelectedValue = 0 Then
                If CDbl(txtValorDoDocumento.Text) = 0 Then
                    ddlMoeda.SelectedValue = 3
                End If
                If CDbl(txtValorDoDocumento.Text) <> 0 Then
                    ddlMoeda.SelectedValue = 1
                End If

                If CDbl(txtValorDoDocumento.Text) <> 0 And CDbl(txtValorEmMoeda.Text) <> 0 Then
                    ddlMoeda.SelectedValue = 0
                End If
            End If

            Sql = " SELECT Indice " & _
                  "   FROM Cotacoes" & _
                  "  WHERE Data_Id = '" & CDate(Data).ToString("yyyy/MM/dd") & "' AND Indexador_Id = " & ddlIndexador.SelectedValue
            Dim dsContacoes As DataSet = Banco.ConsultaDataSet(Sql, "Cotacoes")

            Sql = "Select DataPedido, isnull(IndiceFixado,0) as IndiceFixado From Pedidos where Pedido_id = " & txtPedido.Text
            Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql, "Pedidos")

            If ddlMoeda.SelectedValue > 0 And (DdlProvisoes.SelectedValue = 2 Or DdlProvisoes.SelectedValue = 3) Then
                If ddlMoeda.SelectedValue = 1 Then
                    txtValorEmMoeda.Text = "0"
                Else
                    If (DdlProvisoes.Text = 2 Or DdlProvisoes.Text = 3) Then
                        If txtPedido.Text > 0 Then
                            If dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                                txtValorCobrado.Text = Math.Round(CDbl(txtValorEmMoeda.Text) * dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                            Else
                                txtValorCobrado.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtValorEmMoeda.Text), ddlIndexador.SelectedValue, dsPedido.Tables(0).Rows(0).Item("DataPedido")), 2)
                            End If

                            If CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text) < 0 Then
                                txtAcrescimos.Text = (CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text)) * -1
                            Else
                                txtDeducoes.Text = (CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text))
                            End If
                        Else
                            txtValorDoDocumento.Text = FormatNumber(CDbl(txtValorEmMoeda.Text) * dsContacoes.Tables(0).Rows(0).Item("Indice"), 2)
                        End If
                    End If
                End If
            ElseIf DdlProvisoes.Text = 1 AndAlso ddlMoeda.SelectedValue = 1 Then
                If ddlMoeda.SelectedValue = 1 Then
                    txtValorEmMoeda.Text = "0"
                Else
                    If txtPedido.Text > 0 AndAlso dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                        txtValorEmMoeda.Text = Math.Round(CDbl(txtValorDoDocumento.Text) / dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                    Else
                        txtValorEmMoeda.Text = FormatNumber(CDbl(txtValorDoDocumento.Text) / dsContacoes.Tables(0).Rows(0).Item("Indice"), 2)
                    End If
                End If
            End If
        End If
    End Sub

    Private Function DolarizaBaixa(ByVal Valor As String, ByVal Indexador As String) As String
        Dim Calculo As Decimal
        If Data = "" Or Data Is Nothing Then Data = txtMovimento.Text

        Sqla = " SELECT Indice " & _
                    "   FROM Cotacoes" & _
               "  WHERE Data_Id = '" & CDate(Data).ToString("yyyy/MM/dd") & "' " & vbCrLf & _
               "    AND Indexador_Id = " & Indexador & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
            Calculo = CDec(Valor) / IIf(Dr("Indice") = 0, 1, Dr("Indice"))
            'Calculo = CDec(FormatNumber(Calculo, 2))
        Next

        Return Calculo.ToString("N2")
    End Function

    Function Carteira(ByVal Valor As String) As String
        Dim Conta As String = ""

        If DdlTributos.Text <> "" Then
            Sqla = " SELECT ContaCredito AS Conta " & vbCrLf & _
                   "   FROM Encargos " & vbCrLf & _
                   "  WHERE Encargo_Id = '" & Valor & "'"
        Else
            Sqla = " SELECT ContaClientes AS Conta " & vbCrLf & _
                   "   FROM ComprasXProdutos" & vbCrLf & _
                   "  WHERE Produto_Id = '" & Valor & "'" & vbCrLf
        End If

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
            Conta = Dr("Conta")
        Next
        Return Conta
    End Function

    Sub ConsultaCarteiras(ByVal Carteira As String)
        Sql = "SELECT ContaClientes, ContaDescontos, ContaDeducoes, ContaJuros, ContaAcrescimos" & vbCrLf & _
              "  FROM ComprasXProdutos " & vbCrLf & _
              " WHERE Produto_Id = '" & Carteira & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Carteiras").Tables(0).Rows
            ContaClientes = Dr("ContaClientes")
            ContaDescontos = Dr("ContaDescontos")
            ContaDeducoes = Dr("ContaDeducoes")
            ContaJuros = Dr("ContaJuros")
            ContaAcrescimos = Dr("ContaAcrescimos")
        Next
    End Sub

    Protected Sub DdlEmpresaPagadora_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DdlTiposDePagamentos.SelectedIndex = 0
        CargaBancoRecebedor()
    End Sub

    Function PesoPago(ByVal Pedido As String, ByVal Valor As String, ByVal Historico As String) As String
        Dim Empresa() As String = DdlEmpresaCliente.SelectedValue.ToString.Split("-")
        Dim HistoricoParcial As String = txtHistorico.Text

        Sqla = "SELECT pif.Pedido_Id," & vbCrLf & _
               "       pif.Produto_Id," & vbCrLf & _
               "       Produtos.Nome," & vbCrLf & _
               "       Produtos.Agrupar," & vbCrLf & _
               "       pif.Fixacao_Id," & vbCrLf & _
               "       pif.Quantidade," & vbCrLf & _
               "       isnull(p.PedidoEfetivo,'') as PedidoEfetivo," & vbCrLf & _
               "       ISNULL((SELECT SUM( T1.ValorOficial) AS Oficial" & vbCrLf & _
               "                 FROM VW_PedidosXItensXFixacoesXEncargos AS T1" & vbCrLf & _
               "                WHERE T1.Encargo_Id = 'LIQUIDO'" & vbCrLf & _
               "                  AND T1.Pedido_Id  = pif.Pedido_Id), 0) AS Oficial," & vbCrLf & _
               "       ISNULL((SELECT SUM(T1.ValorMoeda) AS Moeda" & vbCrLf & _
               "                 FROM VW_PedidosXItensXFixacoesXEncargos AS T1" & vbCrLf & _
               "                WHERE T1.Encargo_Id = 'LIQUIDO'" & vbCrLf & _
               "                  AND T1.Pedido_Id = pif.Pedido_Id), 0) AS Moeda" & vbCrLf & _
               "  FROM VW_PedidosXItensXFixacoes pif" & vbCrLf & _
               " INNER JOIN Produtos" & vbCrLf & _
               "    ON pif.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
               " Inner join Pedidos P" & vbCrLf & _
               "    on P.Empresa_id    = pif.empresa_id" & vbCrLf & _
               "   and P.EndEmpresa_id = pif.EndEmpresa_Id" & vbCrLf & _
               "   and P.Pedido_id     = pif.Pedido_id" & vbCrLf & _
               " WHERE P.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf & _
               "   AND pif.Pedido_Id      = " & Pedido

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Pedido").Tables(0).Rows
            If Dr("Agrupar") = "N" And Dr("Oficial") <> 0 Then
                HistoricoParcial = Left(Historico, 21)
                If ddlMoeda.SelectedValue = 1 Then
                    HistoricoParcial &= Convert.ToDecimal(Dr("Quantidade") * Valor / Dr("Oficial")).ToString("N0") & " KGS DE " & Dr("Nome") & " - PEDIDO " & Pedido & IIf(Dr("PedidoEfetivo").ToString.Length = 0, "", ", CN - " & Dr("PedidoEfetivo"))
                Else
                    HistoricoParcial &= Convert.ToDecimal(Dr("Quantidade") * Valor / Dr("Moeda")).ToString("N0") & " KGS DE " & Dr("Nome") & " - PEDIDO " & Pedido & IIf(Dr("PedidoEfetivo").ToString.Length = 0, "", ", CN - " & Dr("PedidoEfetivo"))
                End If
            End If
        Next
        Return HistoricoParcial

    End Function

    Protected Sub ddlCarteiras_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCarteiras.SelectedIndexChanged
        If DdlProvisoes.SelectedIndex = 0 Then
            ddlCarteiras.SelectedIndex = 0
            MsgBox(Me.Page, "Selecione Previsão ou Baixa")
        Else
            Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
            DdlTributos.Parent.Visible = Not objCarteira.isAdiantamento
            txtNumeroAdto.Parent.Visible = objCarteira.isAdiantamento
            txtVencimentoAdto.Parent.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

            txtVencimentoAdto.Enabled = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)
            txtTaxaAdto.Parent.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

            If Not objCarteira.isAdiantamento Then
                CargaTributos()
            End If

            If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento Then
                ConsultaSequenciaDeAdiantamento()
            End If

            If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
                Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                If (strCliente IsNot Nothing AndAlso strCliente.Length > 0) Then
                    Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
                    If objCliente.DesdobrarFornecedor = True Then
                        ucDestinoContabil.Limpar()
                        Dim parameters As New Dictionary(Of String, Object)
                        parameters.Add("tipo", "R")
                        Popup.ConsultaDeDestinoContabil(Me.Page, "objDestinoContabil" & HID.Value)
                        ucDestinoContabil.Carregar(parameters)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ConsultaSequenciaDeAdiantamento()
        Dim emp() As String = DdlEmpresaCliente.SelectedValue.ToString.Split("-")
        Dim NumAdiantamento As Integer

        If IsNumeric(txtRegistro.Text) AndAlso CInt(txtRegistro.Text) > 0 Then
            Sql = "SELECT isnull(A.Adiantamento_id,0) as Adiantamento" & vbCrLf & _
                  "  FROM ContasAReceber cr" & vbCrLf & _
                  "  Left join Adiantamentos A" & vbCrLf & _
                  "    on cr.registro_id = a.titulo " & vbCrLf & _
                  " WHERE cr.Registro_id = " & txtRegistro.Text

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Adto").Tables(0).Rows
                NumAdiantamento = Dr("Adiantamento")
            Next
        End If

        If NumAdiantamento = 0 Then
            Dim num As New [Lib].Negocio.Numerador(emp(0), emp(1), 15)
            NumAdiantamento = num.Sequencia + 1
        End If

        txtNumeroAdto.Text = NumAdiantamento
    End Sub

    Sub CargaTributos()
        Sql = " SELECT Tributo_Id as Codigo, (Encargos.Descricao + ' - ' + Tributo_Id) as Descricao " & vbCrLf & _
              "   FROM CarteirasXTributos " & vbCrLf & _
              "  INNER JOIN Encargos " & vbCrLf &
              "     ON CarteirasXTributos.Tributo_ID = Encargos.Encargo_id " & vbCrLf &
              "  WHERE Carteira_Id = '" & ddlCarteiras.SelectedValue & "'" & vbCrLf &
              "  ORDER BY Tributo_Id " & vbCrLf

        DdlTributos.DataValueField = "Codigo"
        DdlTributos.DataTextField = "Descricao"
        DdlTributos.DataSource = Banco.ConsultaDataSet(Sql, "Tributos")
        DdlTributos.DataBind()

        DdlTributos.Items.Insert(0, "")
        DdlTributos.SelectedIndex = 0
    End Sub

    Function ConsultaTributos(ByVal Tributos As String, Optional ByVal Validacao As Boolean = False) As Boolean
        ContaClientes = ""

        Sql = "SELECT Encargo_Id, ContaCredito " & vbCrLf & _
              "  FROM Encargos " & vbCrLf & _
              " WHERE Encargo_Id  = '" & Tributos & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Encargos")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            If Not Validacao Then
                MsgBox(Me.Page, "Encargo " & Tributos & " não encontrado.")
                Return False
            End If
        Else
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Encargos").Tables(0).Rows
                ContaClientes = Dr("ContaCredito")
            Next
            If ContaClientes.Length = 0 Then
                If Not Validacao Then
                    MsgBox(Me.Page, "Encargo " & Tributos & " sem a conta de Crédito informada.")
                    Return False
                End If
            End If
        End If
        Return True
    End Function

    Function ConsultaTributosCDebito(ByVal Tributos As String, Optional ByVal Validacao As Boolean = False) As Boolean
        ContaClientes = ""
        Sql = "SELECT Encargo_Id, ContaDebito" & vbCrLf & _
              "  FROM Encargos " & vbCrLf & _
              " WHERE Encargo_Id  = '" & Tributos & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Encargos")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            If Not Validacao Then
                MsgBox(Me.Page, "Encargo " & Tributos & " não encontrado.")
                Return False
            End If
        Else
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Encargos").Tables(0).Rows
                ContaClientes = Dr("ContaDebito")
            Next
            If ContaClientes.Length = 0 Then
                If Not Validacao Then
                    MsgBox(Me.Page, "Encargo " & Tributos & " sem a conta de Debito informada.")
                    Return False
                End If
            Else
                Return True
            End If
        End If
    End Function

    Protected Sub RbGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        chkPrevisao.Checked = False
        chkProvisao.Checked = False
    End Sub

    Protected Sub RbAtivo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If chkPrevisao.Visible = True Then
            chkPrevisao.Checked = True
            chkProvisao.Checked = False
        End If
    End Sub

    Protected Sub RbBaixado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        chkPrevisao.Checked = False
        chkProvisao.Checked = False
    End Sub

    Protected Sub RbCancelado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        chkPrevisao.Checked = False
        chkProvisao.Checked = False
    End Sub

    Protected Sub lnkRelatorioCons_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkRelatorioCons.Click
        Try
            If Funcoes.VerificaPermissao("ContasAReceber", "RELATORIO") Then

                Dim crpt As New ReportDocument()

                Try
                    If RbDiaGeral.Checked = True Then
                        DiaGeral()
                    End If
                    If RbFilialDiario.Checked = True Then
                        FilialGeral()
                    End If
                    If RbCarteiraDia.Checked = True Then
                        CarteiraDia()
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                Finally
                    crpt.Close()
                    crpt.Dispose()
                End Try
            Else
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet(ByRef Parametros As String) As DataSet
        Dim ds As New DataSet

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.Text) Then
            Parametros &= "Empresa: " & DdlEmpresaConsultaTitulos.SelectedItem.Text & vbCrLf
        End If

        If txtCodigoClienteConsTitulo.Value.Length > 0 Then
            Parametros &= "Cliente " & IIf(chkClientes.Checked, " - Consolidado:", ":") & txtClienteConsTitulo.Text & vbCrLf
        End If

        Sql = " SELECT Empresas.Reduzido AS ReduzidoEmpresa, CR.Empresa, CR.EndEmpresa, Empresas.Nome AS NomeEmpresa," & vbCrLf & _
              "        Empresas.Nome AS NomeEmpresa, Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, " & vbCrLf & _
              "        CONVERT(NVARCHAR,CR.Registro_Id) + ' ' +" & vbCrLf & _
              "        CASE" & vbCrLf & _
              "            WHEN ISNULL(CR.Moeda,1) = 1 " & vbCrLf & _
              "                THEN 'R$'" & vbCrLf & _
              "                ELSE 'U$'" & vbCrLf & _
              "        END Registro," & vbCrLf & _
              "        '' as Faturamento," & vbCrLf & _
              "        '' as Lote, " & vbCrLf & _
              "        0 as LoteTotal," & vbCrLf & _
              "        0 as LoteEntregue," & vbCrLf & _
              "        CR.Pedido," & vbCrLf & _
              "        CR.Cliente, Clientes.Nome AS NomeCliente," & vbCrLf & _
              "        CR.Movimento, " & vbCrLf & _
              "        CR.Prorrogacao AS Vencimento, " & vbCrLf & _
              "        CR.Baixa, " & vbCrLf & _
              "        CR.Carteira, " & vbCrLf & _
              "        CR.Provisao, " & vbCrLf & _
              "        Carteira.Descricao AS NomeCarteira, " & vbCrLf & _
              "        CR.Historico " & IIf(chkObservacao.Checked, "+ ' / OBS: ' + cast(CR.Observacoes as varchar) as Historico,", ",") & vbCrLf & _
              "        CR.solicitacao, " & vbCrLf & _
              "        CASE " & vbCrLf & _
              "            WHEN CR.Provisao = 1 " & vbCrLf & _
              "			        THEN CR.ValorDoDocumento" & vbCrLf & _
              "                 ELSE " & vbCrLf & _
              "                     CASE " & vbCrLf & _
              "                          WHEN CR.Moeda <> 1 " & vbCrLf & _
              "                               THEN convert(numeric(18,2), CR.MoedaValorDoDocumento * case when CR.Indexador = 99 then p.indiceFixado else Cotacoes.indice end) " & vbCrLf & _
              "                               ELSE CR.ValorDoDocumento " & vbCrLf & _
              "                           END " & vbCrLf & _
              "        END ValorLiquido , " & vbCrLf & _
              "        CASE  " & vbCrLf & _
              "            WHEN CR.Moeda = 1 " & vbCrLf & _
              "            THEN 0" & vbCrLf & _
              "	      	 ELSE CASE " & vbCrLf & _
              "	      	         WHEN CR.Provisao = 1 " & vbCrLf & _
              "	      			      THEN CR.MoedaValorDoDocumento " & vbCrLf & _
              "	      				  ELSE convert(numeric(18,2),CR.ValorDoDocumento / case when CR.Indexador = 99 then p.indiceFixado else Cotacoes.indice end) " & vbCrLf & _
              "                 END " & vbCrLf & _
              "        END MoedaValorLiquido," & vbCrLf & _
              "        (SELECT ISNULL(P.PedidoEfetivo,0) " & vbCrLf & _
              "            FROM Pedidos P " & vbCrLf & _
              "           WHERE P.Pedido_id     = CR.Pedido" & vbCrLf & _
              "             AND P.Empresa_Id    = CR.Empresa " & vbCrLf & _
              "             AND P.EndEmpresa_Id = CR.EndEmpresa " & vbCrLf & _
              "        ) PedidoEfetivo, " & vbCrLf & _
              "       'R' AS Tipo " & vbCrLf & _
              "   FROM  ContasAReceber AS CR" & vbCrLf & _
              "   LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf & _
              "     ON CR.Empresa      = NFXT.Empresa_Id" & vbCrLf & _
              "    AND CR.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf & _
              "    AND CR.Registro_Id = NFXT.Titulo_Id" & vbCrLf & _
              "   LEFT JOIN FaturaDeFreteXTitulo FFXT" & vbCrLf & _
              "     ON CR.Registro_Id = FFXT.Titulo_Id" & vbCrLf & _
              "   LEFT JOIN FaturasDeFretesXItens FFXI" & vbCrLf & _
              "     ON FFXI.EmpresaPagadora_Id    = FFXT.Empresa_Id" & vbCrLf & _
              "    AND FFXI.EndEmpresaPagadora_Id = FFXT.EndEmpresa_Id" & vbCrLf & _
              "    AND FFXI.Fatura_Id             = FFXT.Fatura_Id" & vbCrLf & _
              "   LEFT JOIN Cotacoes " & vbCrLf & _
              "     ON Cotacoes.Data_id      = CR.Prorrogacao " & vbCrLf & _
              "    AND Cotacoes.Indexador_Id = CR.Indexador " & vbCrLf & _
              "  INNER JOIN Clientes AS Empresas " & vbCrLf & _
              "     ON CR.Empresa    = Empresas.Cliente_Id " & vbCrLf & _
              "    AND CR.EndEmpresa = Empresas.Endereco_Id " & vbCrLf & _
              "  INNER JOIN Clientes " & vbCrLf & _
              "     ON CR.Cliente    = Clientes.Cliente_Id " & vbCrLf & _
              "    AND CR.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
              "   LEFT OUTER JOIN ComprasXProdutos AS Carteiras " & vbCrLf & _
              "     ON CR.Carteira = Carteiras.Produto_Id" & vbCrLf & _
              "   LEFT OUTER JOIN Carteira " & vbCrLf & _
              "     ON CR.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf & _
              "    AND (CR.Situacao = 1) " & vbCrLf & _
              "    AND (CR.Grupado <> 'M') " & vbCrLf & _
              "    AND (CR.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf & _
              "    AND (CR.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf & _
              "   LEFT JOIN Pedidos AS P " & vbCrLf & _
              "     ON P.Empresa_id    = CR.EmpresaPedido" & vbCrLf & _
              "    AND P.EndEmpresa_Id = CR.EndEmpresaPedido " & vbCrLf & _
              "    AND P.Pedido_id     = CR.Pedido " & vbCrLf

        If RbCancelado.Checked Then
            Sql &= " WHERE CR.Situacao <> 1 " & vbCrLf
        Else
            Sql &= " WHERE CR.Situacao = 1 " & vbCrLf
        End If

        If RbAtivo.Checked Then
            If chkPrevisao.Visible = True And chkProvisao.Visible = True Then
                If chkPrevisao.Checked = True AndAlso chkProvisao.Checked = True Then
                    Sql &= "    AND (CR.Provisao = 2 OR CR.Provisao = 3) " & vbCrLf
                ElseIf chkPrevisao.Checked = True Then
                    Sql &= "    AND CR.Provisao = 2 " & vbCrLf
                ElseIf chkProvisao.Checked = True Then
                    Sql &= "    AND CR.Provisao = 3 " & vbCrLf
                Else
                    Sql &= "    AND CR.Provisao = 2 " & vbCrLf
                End If
            Else
                Sql &= "    AND CR.Provisao = 2 " & vbCrLf
            End If
        End If

        If RbBaixado.Checked Then
            Sql &= "    AND CR.Provisao = 1 " & vbCrLf
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            Sql &= "    AND CR.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        campo = Cliente.Split("-")
        If campo(0) <> "" Then
            Sql &= "    AND CR.Empresa = '" & campo(0) & "'" & vbCrLf   'Empresa
            Sql &= "    AND CR.EndEmpresa = " & campo(1) & vbCrLf    'Endereco da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= "    AND (ISNULL(NFXT.Nota_Id,0) = " & txtNumNota.Text & " OR ISNULL(FFXI.Nota_Id,0) = " & txtNumNota.Text & ")"
        Else
            Cliente = txtCodigoClienteConsTitulo.Value
            campo = Cliente.Split("-")
            If Not String.IsNullOrWhiteSpace(campo(0)) Then
                If chkClientes.Checked Then
                    Sql &= "    AND left(CR.Cliente, 8) = '" & Left(campo(0), 8) & "'" & vbCrLf
                Else
                    Sql &= "    AND CR.Cliente = '" & campo(0) & "'" & vbCrLf & _
                           "    AND CR.EndCliente = " & campo(1) & vbCrLf
                End If
            End If

            If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
                Sql &= "    AND CR.Pedido= '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
                Sql &= "    AND CR.Pedido= '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
            End If

            If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
                Sql &= "    AND CR.Prorrogacao between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
            End If
        End If

        Sql &= "  ORDER BY CR.Prorrogacao, Clientes.Nome, CR.Registro_Id" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        Return ds
    End Function

    Private Sub DiaGeral()
        Dim ds As New DataSet
        Dim Parametros As String = ""
        Dim crystal As String = ""

        ds = getDataSet(Parametros)

        If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            crystal = "Cr_Titulos"
        Else
            crystal = "Cr_TitulosPorEmpresa"
        End If

        Dim parameters = New Dictionary(Of String, Object)()
        parameters.Add("Relatorio", "Relação de Títulos A Receber.")
        Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))
        parameters.Add("EmpresaNome", objEmpresa.Nome)
        parameters.Add("EmpresaCidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
        parameters.Add("EmpresaCodigo", Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
        parameters.Add("UnidadeDeNegocio", Parametros)
        parameters.Add("TipoDaCarteira", "Carteira")

        Funcoes.BindReport(Me.Page, ds, crystal, eExportType.PDF, parameters)
    End Sub

    Sub FilialGeral()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        Dim Valdia As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        Dim Empresa As String

        sql = " SELECT CR.Registro_Id AS Registro, " & vbCrLf & _
              "        convert(varchar(10),CR.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              "        CR.Prorrogacao,103) as Posterga, Clientes.Nome AS Cliente, Historico, " & vbCrLf & _
              "        ISNULL(CR.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              "        CR.ValorLiquido AS Valor, " & vbCrLf & _
              "        UsuarioLiberacao as Liberado, " & vbCrLf & _
              "        CR.Carteira As Carteira, " & vbCrLf & _
              "        ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              "        CR.Situacao as Situacao, " & vbCrLf & _
              "        CR.Empresa as Empresa, " & vbCrLf & _
              "        Empresa.Reduzido as Reduzido, " & vbCrLf & _
              "        CR.UsuarioBaixa as UsuarioBaixa, " & vbCrLf & _
              "        CR.Pedido as Pedido" & vbCrLf & _
              "   FROM ContasAReceber CR " & vbCrLf & _
              "  INNER JOIN Clientes" & vbCrLf & _
              "     ON CR.Cliente = Clientes.Cliente_Id " & vbCrLf & _
              "    AND CR.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
              "   LEFT JOIN ComprasXProdutos" & vbCrLf & _
              "     ON CR.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              "  INNER JOIN Clientes as Empresa" & vbCrLf & _
              "     ON CR.Empresa = Empresa.Cliente_Id " & vbCrLf & _
              "    AND CR.EndEmpresa = Empresa.Endereco_Id" & vbCrLf & _
              "  WHERE CR.Provisao <> 1 " & vbCrLf & _
              "    AND CR.Situacao = 1" & vbCrLf

        lista = "Geral"

        If RbAtivo.Checked = True Then
            sql &= "    AND CR.usuariobaixa = '' " & vbCrLf
            lista = "Ativos" & vbCrLf
        End If

        If RbBaixado.Checked = True Then
            sql &= "    AND CR.usuariobaixa <> '' " & vbCrLf
            lista = "Baixados" & vbCrLf
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= "    AND CR.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= "    AND CR.Empresa = '" & Campo(0) & "'" & vbCrLf  'Empresa
            sql &= "    AND CR.EndEmpresa = " & Campo(1) & vbCrLf     'Endereco da Empresa
        End If

        Cliente = txtCodigoClienteConsTitulo.Value
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= "    AND CR.Cliente = '" & Campo(0) & "'" & vbCrLf  'Cliente
            sql &= "    AND CR.EndCliente = " & Campo(1) & vbCrLf      'Cliente da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
            sql &= "    AND CR.Pedido = '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
        End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= "    AND Vencimento between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
        End If

        sql &= "    ORDER BY CR.Empresa, CR.Vencimento " & vbCrLf

        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf
        linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
        linha &= "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf
        linha &= "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        '-----------------
        'Cabeçalho Padrao
        '-----------------
        linha &= "<TR>"
        linha &= "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>"
        linha &= "</TR>"
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>"
            linha &= "</TR>"
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>"
            linha &= "</TR>"
        End If

        Cliente = txtCodigoClienteConsTitulo.Value
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & txtCodigoClienteConsTitulo.Value & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>"
            linha &= "</TR>"
        End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>"
        linha &= "<TD >" & " Tipo de relatorio Filial(Empresa)Geral - Totalizacao por Filial - Registros impressos : " & lista & " </TD>"
        linha &= "</TR>"
        linha &= "<table width= '370' cellpadding='0' cellspacing='0' Border=0>"
        linha &= "<TR>"
        linha &= "<TD><B>Registro</B></TD>"
        linha &= "<TD><B>Cliente/Fornecedor</B></TD>"
        linha &= "<TD><B>Historico</B></TD>"
        linha &= "<TD><B>Receber R$</B></TD>"
        linha &= "<TD><B>Receber US$</B></TD>"
        linha &= "<TD><B>Empresa</B></TD>"
        linha &= "<TD><B>Vencimento Original</B></TD>"
        linha &= "<TD><B>Prorrogacao</B></TD>"
        linha &= "<TD><B>Carteira</B></TD>"
        linha &= "<TD><B>Situacao</B></TD>"
        linha &= "</TR>"
        ''
        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        Empresa = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If Empresa <> "" And Empresa <> dr("Empresa") Then
                    linha &= "<TR>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & " Total Filial: " & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "</TR>"
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")
                Empresa = dr("Empresa")
                Datvenctr = dr("Vencimento")
                linha &= "<TR>"
                linha &= "<TD>" & dr("Registro") & "</TD>"
                linha &= "<TD>" & dr("Cliente") & "</TD>"
                linha &= "<TD>" & dr("Historico") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>"
                linha &= "<TD>" & dr("Reduzido") & "</TD>"
                linha &= "<TD>" & dr("Vencimento") & "</TD>"
                linha &= "<TD>" & dr("Posterga") & "</TD>"
                linha &= "<TD>" & dr("DescricaoCarteira") & "</TD>"
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If
        linha &= "<TR>>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total Filial: " & "</B></TD>"
        linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        linha &= "<TR>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total neste processamento: " & "</B></TD>"
        linha &= "<TD><B>" & Valtot.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        If contreg = 0 Then
            MsgBox(Me.Page, "Nao existem registros corretos para o periodo !!!!")
            txtPeriodoFinalConsultaTitulos.Focus()
        Else
            MsgBox(Me.Page, "Movimento com registros processados !!!")
            txtPeriodoFinalConsultaTitulos.Focus()
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
        Finally
            strm.Close()
        End Try
    End Sub

    Sub CarteiraDia()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        Dim Valdia As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        Dim Empresa As String
        Dim carteira As String

        sql = "SELECT CR.Registro_Id AS Registro, " & vbCrLf & _
              "       CONVERT(VARCHAR(10),CR.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              "       CR.Prorrogacao,103) as Posterga, Cli.Nome AS Cliente, Historico, " & vbCrLf & _
              "       ISNULL(CR.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              "       CR.ValorLiquido AS Valor, " & vbCrLf & _
              "       UsuarioLiberacao as Liberado, " & vbCrLf & _
              "       CR.Carteira As Carteira, " & vbCrLf & _
              "       ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              "       CR.Situacao as Situacao, " & vbCrLf & _
              "       CR.Empresa as Empresa, " & vbCrLf & _
              "       Empresa.Reduzido as Reduzido, " & vbCrLf & _
              "       CR.UsuarioBaixa as UsuarioBaixa, " & vbCrLf & _
              "       CR.Pedido as Pedido" & vbCrLf & _
              "  FROM ContasAReceber CR " & vbCrLf & _
              " INNER JOIN Clientes Cli" & vbCrLf & _
              "    ON CR.Cliente = Cli.Cliente_Id " & vbCrLf & _
              "   AND CR.EndCliente = Cli.Endereco_Id" & vbCrLf & _
              "  LEFT JOIN ComprasXProdutos" & vbCrLf & _
              "    ON CR.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              " INNER JOIN Clientes as Empresa" & vbCrLf & _
              "    ON CR.Empresa = Empresa.Cliente_Id " & vbCrLf & _
              "   AND CR.EndEmpresa = Empresa.Endereco_Id" & vbCrLf & _
              " WHERE CR.Provisao <> 1 " & vbCrLf & _
              "   AND  CR.Situacao = 1" & vbCrLf

        lista = "Geral"

        If RbAtivo.Checked = True Then
            sql &= "   AND CR.usuariobaixa = '' " & vbCrLf
            lista = "Ativos"
        End If

        If RbBaixado.Checked = True Then
            sql &= "   AND CR.usuariobaixa <> '' " & vbCrLf
            lista = "Baixados"
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= "   AND CR.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= "   AND CR.Empresa = '" & Campo(0) & "'" & vbCrLf  'Empresa
            sql &= "   AND CR.EndEmpresa = " & Campo(1) & vbCrLf     'Endereco da Empresa
        End If

        Cliente = txtCodigoClienteConsTitulo.Value
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= "   AND CR.Cliente = '" & Campo(0) & "'" & vbCrLf   'Cliente
            sql &= "   AND CR.EndCliente = " & Campo(1) & vbCrLf     'Cliente da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
            sql &= "   AND CR.Pedido = '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
        End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= "   AND Vencimento between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
        End If

        sql &= " ORDER BY CR.Carteira, CR.Vencimento " & vbCrLf

        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf
        linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
        linha &= "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf
        linha &= "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        '-----------------
        'Cabeçalho Padrao
        '-----------------
        linha &= "<TR>"
        linha &= "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>"
        linha &= "</TR>"
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>"
            linha &= "</TR>"
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>"
            linha &= "</TR>"
        End If

        Cliente = txtCodigoClienteConsTitulo.Value
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & txtCodigoClienteConsTitulo.Value & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>"
            linha &= "</TR>"
        End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>"
        linha &= "<TD >" & " Tipo de relatorio Filial(Empresa)Geral - Totalizacao por Filial - Registros impressos : " & lista & " </TD>"
        linha &= "</TR>"
        linha &= "<table width= '370' cellpadding='0' cellspacing='0' Border=0>"
        linha &= "<TR>"
        linha &= "<TD ><B>Registro</B></TD>"
        linha &= "<TD ><B>Cliente/Fornecedor</B></TD>"
        linha &= "<TD ><B>Historico</B></TD>"
        linha &= "<TD ><B>Receber R$</B></TD>"
        linha &= "<TD ><B>Receber US$</B></TD>"
        linha &= "<TD ><B>Empresa</B></TD>"
        linha &= "<TD ><B>Vencimento Original</B></TD>"
        linha &= "<TD ><B>Prorrogacao</B></TD>"
        linha &= "<TD ><B>Carteira<B></TD>"
        linha &= "<TD ><B>Situacao</B></TD>"
        linha &= "</TR>"
        ''
        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        Empresa = ""
        carteira = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If carteira <> "" And carteira <> dr("DescricaoCarteira") Then
                    linha &= "<TR>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & " Total Da Carteira: " & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "</TR>"
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")
                Empresa = dr("Empresa")
                Datvenctr = dr("Vencimento")
                carteira = dr("DescricaoCarteira")
                linha &= "<TR>"
                linha &= "<TD>" & dr("Registro") & "</TD>"
                linha &= "<TD>" & dr("Cliente") & "</TD>"
                linha &= "<TD>" & dr("Historico") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>"
                linha &= "<TD>" & dr("Reduzido") & "</TD>"
                linha &= "<TD>" & dr("Vencimento") & "</TD>"
                linha &= "<TD>" & dr("Posterga") & "</TD>"
                linha &= "<TD>" & dr("Carteira") & " - " & dr("DescricaoCarteira") & "</TD>"
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If
        linha &= "<TR>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total Da Carteira: " & "</B></TD>"
        linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        linha &= "<TR>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total neste processamento: " & "</B></TD>"
        linha &= "<TD><B>" & Valtot.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        If contreg = 0 Then
            MsgBox(Me.Page, "Nao existem registros corretos para o periodo !!!!")
            txtPeriodoFinalConsultaTitulos.Focus()
        Else
            MsgBox(Me.Page, "Movimento com registros processados !!!")
            txtPeriodoFinalConsultaTitulos.Focus()
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
        Finally
            strm.Close()
        End Try
        '' rotina utilizada nas procuracoes fim - html 
        '' rotina de geracao do relatorio . 
    End Sub

    Private Sub EmitirRecibo()

        Try



            Dim Tdigito As String = ""
            Dim TDigitoAgencia As String = ""
            Dim xextenso As String = ""
            Dim yextenso As String = ""
            Dim row As DataRow
            Dim mes As Integer
            Dim MESX As String = ""
            Dim Historico As String = ""
            Dim RegistroI As String = ""
            Dim RegistroS As String = ""
            Dim Empresa As String = ""
            Dim EndEmpresa As String = ""
            Dim ENome As String = ""
            Dim EEndereco As String = ""
            Dim ECep As String = ""
            Dim ECidade As String = ""
            Dim EEstado As String = ""
            Dim ECnpj As String = ""
            Dim EInscricao As String = ""
            Dim Efone As String = ""
            Dim EBairro As String = ""
            Dim EComplemento As String = ""
            Dim ENumero As Integer
            Dim Cliente As String = ""
            Dim EndCliente As String = ""
            Dim CNome As String = ""
            Dim CEndereco As String = ""
            Dim CCep As String = ""
            Dim CCidade As String = ""
            Dim CEstado As String = ""
            Dim CCnpj As String = ""
            Dim CInscricao As String = ""
            Dim Cfone As String = ""
            Dim CBairro As String = ""
            Dim CComplemento As String = ""
            Dim CNumero As Integer
            Dim CBancoCliente As String = ""
            Dim CAgenciaCliente As String = ""
            Dim CDigitoAgenciaCliente As String = ""
            Dim CCcontaCliente As String = ""
            Dim CDigitoContaCliente As String = ""
            Dim dsRecibo As New DataSet
            Dim dtRecibo As DataTable

            dtRecibo = New DataTable("ReciboAPagar")
            '' campos da empresa 
            dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
            ' ''campos do cliente
            dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
            '' campos to titulo 
            dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
            dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
            dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))
            dsRecibo.Tables.Add(dtRecibo)
            Dim ValorCobrado As Decimal
            Dim ValorDoDocumento As Decimal
            Dim Juros As Decimal
            Dim Acrescimos As Decimal
            Dim Descontos As Decimal
            Dim deducoes As Decimal
            Dim TipoPagto As Integer
            Dim FormaDePagamento As String = ""
            Dim DataBaixa As String = ""
            Dim dataVencimento As String = ""
            Dim dr As DataRow

            '' Definicao de variaveis fim. 
            '' Lendo campos inicio
            Registro = txtRegistro.Text
            ''txtRegistro.Text = Registro
            '' Lendo titulo - inicio .'. 
            Sql = "SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf & _
                  "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf & _
                  "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf & _
                  "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido,0) AS Pedido, isnull(PedidoFixacao,0) AS PedidoFixacao, isnull(Procuracao,0) AS Procuracao, ValorDoDocumento, " & vbCrLf & _
                  "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf & _
                  "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf & _
                  "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, (case when len(convert(nvarchar(4000),Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),Observacoes ) else Historico end) as Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf & _
                  "       Solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf & _
                  "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, ISNULL(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,T.Descricao " & vbCrLf & _
                  "  FROM ContasAReceber " & vbCrLf & _
                  " INNER JOIN TIPOSDEPAGAMENTOS AS T ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf & _
                  " WHERE Registro_Id = " & Registro


            Dim dsCR As DataSet = Banco.ConsultaDataSet(Sql, "ContasAReceber")
            If dsCR.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Tipo de pagamento do Título não foi definido.")
                Exit Sub
            End If

            For Each Dr1 As DataRow In dsCR.Tables(0).Rows
                DataBaixa = CDate(Dr1("Baixa")).ToString("dd/MM/yyyy")
                dataVencimento = CDate(Dr1("Vencimento")).ToString("dd/MM/yyyy")
                Empresa = Dr1("Empresa")
                EndEmpresa = Dr1("EndEmpresa")
                Cliente = Dr1("Cliente")
                EndCliente = Dr1("EndCliente")
                ValorCobrado = 0
                ValorDoDocumento = Dr1("ValorDoDocumento")
                Juros = Dr1("Juros")
                Acrescimos = Dr1("Acrescimos")
                Descontos = Dr1("Descontos")
                deducoes = Dr1("Deducoes")
                If Dr1("Observacoes").ToString.Length > 0 Then
                    Historico = Funcoes.EliminarCaracteresEspeciais(Dr1("Historico")) & ". " & Funcoes.EliminarCaracteresEspeciais(Dr1("Observacoes"))
                Else
                    Historico = Funcoes.EliminarCaracteresEspeciais(Dr1("Historico"))
                End If
                CBancoCliente = Dr1("BancoCliente")
                CAgenciaCliente = Dr1("AgenciaCliente")
                CDigitoAgenciaCliente = Dr1("DigitoAgenciaCliente")
                CCcontaCliente = Dr1("ContaCliente")
                CDigitoContaCliente = Dr1("DigitoContaCliente")
                ValorCobrado = ValorDoDocumento + Juros + Acrescimos - Descontos - deducoes
                TipoPagto = Dr1("TipoPagto")
                FormaDePagamento = Dr1("Descricao")
                Tdigito = Dr1("DigitoContaCliente")
                TDigitoAgencia = Dr1("DigitoAgenciaCliente")
            Next
            '' Lendo titulo fim .'. 
            '' Consultado empresa 
            Sql = "  SELECT Cli.Cliente_Id ," & vbCrLf & _
                  "         Cli.Nome, Cli.Cidade," & vbCrLf & _
                  "         Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                  "         Cli.Endereco , Cli.Cep," & vbCrLf & _
                  "         Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                  "         Cli.Bairro, Cli.Complemento," & vbCrLf & _
                  "         Cli.Numero " & vbCrLf & _
                  "    FROM Clientes Cli" & vbCrLf & _
                  "   WHERE Cli.Cliente_Id = '" & Empresa & "'" & vbCrLf
            DS = Banco.ConsultaDataSet(Sql, "Clientes")
            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    ENome = dr("Nome")
                    EEndereco = dr("Endereco")
                    ECep = dr("Cep")
                    ECidade = dr("Cidade")
                    EEstado = dr("Estado")
                    ECnpj = dr("Cliente_id")
                    EInscricao = dr("Inscricao")
                    Efone = dr("Telefone")
                    EBairro = dr("Bairro")
                    EComplemento = dr("Complemento")
                    ENumero = dr("Numero")
                    Exit For
                Next
            End If
            '' Consultando Empresa - fim 
            ''**************************************************************
            '' Consultado Cliente 
            Sql = " SELECT Cli.Cliente_Id ," & vbCrLf & _
                  "        Cli.Nome, Cli.Cidade," & vbCrLf & _
                  "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                  "        Cli.Endereco , Cli.Cep," & vbCrLf & _
                  "        Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                  "        Cli.Bairro, Cli.Complemento," & vbCrLf & _
                  "        Cli.Numero " & vbCrLf & _
                  "   FROM Clientes Cli" & vbCrLf & _
                  "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf & _
                  "    AND Cli.Endereco_id = " & EndCliente & vbCrLf

            DS = Banco.ConsultaDataSet(Sql, "Clientes")
            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    CNome = dr("Nome")
                    CEndereco = dr("Endereco")
                    CCep = dr("Cep")
                    CCidade = dr("Cidade")
                    CEstado = dr("Estado")
                    CCnpj = dr("Cliente_id")
                    CInscricao = dr("Inscricao")
                    Cfone = dr("Telefone")
                    CBairro = dr("Bairro")
                    CComplemento = dr("Complemento")
                    CNumero = dr("Numero")
                    Exit For
                Next
            End If
            ' Cria Data Sete que vai ser utilizado no relatorio
            '' Consultado Cliente fim 
            '' Lendo campos fim . 
            ' Move campos para o Data Set.
            ' Move campos da Empresa
            row = dtRecibo.NewRow()
            row("ENome") = ENome
            row("EEndereco") = EEndereco
            row("ECep") = ECep
            row("ECidade") = ECidade
            row("EEstado") = EEstado
            row("ECnpj") = Funcoes.FormatarCpfCnpj(ECnpj)
            row("EInscricao") = EInscricao
            row("EFone") = Efone
            row("EBairro") = EBairro
            row("EComplemento") = EComplemento
            row("ENumero") = ENumero
            '' Move campos do Cliente / fornecedor
            row("CNome") = CNome
            row("CEndereco") = CEndereco
            row("CCep") = CCep
            row("CCidade") = CCidade
            row("CEstado") = CEstado
            row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
            row("CInscricao") = CInscricao
            row("CFone") = Cfone
            row("CBairro") = CBairro
            row("CComplemento") = CComplemento
            row("CNumero") = CNumero
            '' Move campos do Titulo
            row("Tnumtit") = Registro
            '' feito
            row("TValor") = ValorCobrado

            Dim valcobradostr As String
            valcobradostr = CStr(ValorCobrado)
            txtValorCobrado.Text = ValorCobrado

            Valor = Replace(txtValorCobrado.Text, ".", "")

            ''* Rotina de extenso inicio
            yextenso = "("
            yextenso &= UCase(Funcoes.Extenso(txtValorCobrado.Text(), "Real", "Reais"))
            yextenso &= " *"
            xextenso = yextenso
            For j = 1 To (120 - Len(xextenso))
                xextenso &= " *"
            Next
            xextenso &= ")"
            row("TExtenso") = xextenso
            ''* Rotina de extenso fim 
            row("THistorico") = Historico
            row("TDia") = Day(DataBaixa)
            row("TAno") = Year(DataBaixa)

            row("TMes") = Month(DataBaixa)
            ''* Rotina do Mes inicio
            mes = Month(DataBaixa)
            If mes = 1 Then
                MESX = "JANEIRO"
            End If

            If mes = 2 Then
                MESX = "FEVEREIRO"
            End If

            If mes = 3 Then
                MESX = "MARCO"
            End If

            If mes = 4 Then
                MESX = "ABRIL"
            End If

            If mes = 5 Then
                MESX = "MAIO"
            End If

            If mes = 6 Then
                MESX = "JUNHO"
            End If

            If mes = 7 Then
                MESX = "JULHO"
            End If

            If mes = 8 Then
                MESX = "AGOSTO"
            End If

            If mes = 9 Then
                MESX = "SETEMBRO"
            End If

            If mes = 10 Then
                MESX = "OUTUBRO"
            End If

            If mes = 11 Then
                MESX = "NOVEMBRO"
            End If

            If mes = 12 Then
                MESX = "DEZEMBRO"
            End If

            row("TMes") = MESX
            ''* rotina do mes fim 


            row("TFormaPagto") = FormaDePagamento
            row("TBanco") = CBancoCliente
            row("TAgencia") = CAgenciaCliente
            row("TConta") = CCcontaCliente
            row("TVencimento") = dataVencimento
            row("TBaixa") = DataBaixa

            '' calculo de registro . 
            RegistroI = Registro
            RegistroS = " "
            RegistroS = Trim(RegistroS)
            If Len(RegistroS) < 6 Then
                For k = 1 To (6 - Len(Trim(RegistroI)))
                    RegistroS &= "0"
                Next
            End If
            row("Trecibo") = Trim(RegistroS) & (Trim(RegistroI))
            row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))
            row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))
            row("TValorDoDocumento") = txtValorDoDocumento.Text
            row("TDescontos") = txtDescontos.Text
            row("TDeducoes") = txtDeducoes.Text
            row("TJuros") = txtJuros.Text
            row("TAcrescimos") = txtAcrescimos.Text
            row("TDigito") = Tdigito
            row("TDigitoAgencia") = TDigitoAgencia
            dtRecibo.Rows.Add(row)

            Dim param As New Dictionary(Of String, Object)
            param.Add("XNome", ENome)

            Funcoes.BindReport(Me.Page, dsRecibo, "Cr_ReciboReceber", eExportType.PDF, param)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEmiteReciboGeral_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim lstSelecionados As List(Of String) = GridConsultaTitulos.GetSelectedValues("chkSelecionado")
        If lstSelecionados Is Nothing AndAlso lstSelecionados.Count <= 0 Then
            MsgBox(Me.Page, "Selecione pelo menos um título para emissão do Recibo.")
            Exit Sub
        End If

        '' rotina de imprimir recibo por registros checados.
        Dim xextenso As String = ""
        Dim yextenso As String = ""
        Dim dsEmitir As New DataSet
        'Dim dtEmitir As DataTable
        Dim row As DataRow
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim mes As Integer
        Dim MESX As String = ""
        'Dim NumeroDoCheque As Integer
        'Dim NumeroDoChequeX As String
        Dim Historico As String = ""
        Dim RegistroI As String = ""
        Dim RegistroS As String = ""
        Dim SqlArray As New ArrayList
        'Dim Sqlupdate As String
        'Dim Campo() As String
        'Dim Mensagem As String
        ''*Definicao de campos da empresa
        Dim Empresa As String = ""
        Dim EndEmpresa As String = ""
        Dim ENome As String = ""
        Dim EEndereco As String = ""
        Dim ECep As String = ""
        Dim ECidade As String = ""
        Dim EEstado As String = ""
        Dim ECnpj As String = ""
        Dim EInscricao As String = ""
        Dim Efone As String = ""
        Dim EBairro As String = ""
        Dim EComplemento As String = ""
        Dim ENumero As Integer
        ''*Definicao de campos da empresa Fim 
        ''* DAdos do cliente inicio
        Dim Cliente As String = ""
        Dim EndCliente As String = ""
        Dim CNome As String = ""
        Dim CEndereco As String = ""
        Dim CCep As String = ""
        Dim CCidade As String = ""
        Dim CEstado As String = ""
        Dim CCnpj As String = ""
        Dim CInscricao As String = ""
        Dim Cfone As String = ""
        Dim CBairro As String = ""
        Dim CComplemento As String = ""
        Dim CNumero As Integer

        Dim CBancoCliente As String = ""
        Dim CAgenciaCliente As String = ""
        Dim CDigitoAgenciaCliente As String = ""
        Dim CCcontaCliente As String = ""
        Dim CDigitoContaCliente As String = ""

        ''* Dados do Cliente fim 
        ''* Campos do DAta Set + titulos inicio
        'Dim myDataRow As DataRow
        '' Cria data Set que vai ser utilizado no relatorio 
        Dim dsRecibo As New DataSet
        Dim dtRecibo As DataTable
        ''Dim row As DataRow
        ''Dim RegistroI As String
        ''Dim RegistroS As String
        dtRecibo = New DataTable("ReciboAPagar")
        '' campos da empresa 
        dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
        ' ''campos do cliente
        dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
        '' campos to titulo 
        dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
        dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
        dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
        dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
        dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))

        dsRecibo.Tables.Add(dtRecibo)
        Dim ValorCobrado As Decimal
        Dim ValorDoDocumento As Decimal
        Dim Juros As Decimal
        Dim Acrescimos As Decimal
        Dim Descontos As Decimal
        Dim deducoes As Decimal
        Dim TipoPagto As Integer
        Dim FormaDePagamento As String = ""
        '' Campos novos do cheque 05/05/2010
        Dim TvalorDoDocumento As Decimal
        Dim Tdescontos As Decimal
        Dim Tdeducoes As Decimal
        Dim TJuros As Decimal
        Dim TAcrescimos As Decimal
        Dim Tdigito As String = ""
        'Dim TnumeroDoCheque As Integer
        '' Campos novos do cheque 05/05/2010

        ''* Campos do Data Set + titulos fim.
        Dim dsCR As DataSet
        While i < GridConsultaTitulos.Rows.Count
            If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then

                '' Rotina para leitura de registro buscando a data da baixa. 
                Dim DataBaixa As String = ""
                Dim dataVencimento As String = ""
                Registro = GridConsultaTitulos.Rows(i).Cells(2).Text()
                Sql = "SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf & _
                      "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf & _
                      "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf & _
                      "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido,0) AS Pedido, isnull(PedidoFixacao,0) AS PedidoFixacao, isnull(Procuracao,0) AS Procuracao, ValorDoDocumento, " & vbCrLf & _
                      "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf & _
                      "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf & _
                      "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, (case when len(convert(nvarchar(4000),Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),Observacoes ) else Historico end) as Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf & _
                      "       solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf & _
                      "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, isnull(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,T.Descricao " & vbCrLf & _
                      "  FROM ContasAReceber " & vbCrLf & _
                      " INNER JOIN TIPOSDEPAGAMENTOS AS T ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf & _
                      " WHERE Registro_Id = " & Registro & vbCrLf


                dsCR = Banco.ConsultaDataSet(Sql, "ContasAReceber")
                If dsCR.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Tipo de pagamento do Título não foi definido.")
                    Exit Sub
                End If

                For Each Dr1 As DataRow In dsCR.Tables(0).Rows
                    DataBaixa = CDate(Dr1("Baixa")).ToString("dd/MM/yyyy")
                    dataVencimento = CDate(Dr1("Prorrogacao")).ToString("dd/MM/yyyy")
                    Empresa = Dr1("Empresa")
                    EndEmpresa = Dr1("EndEmpresa")
                    Cliente = Dr1("Cliente")
                    EndCliente = Dr1("EndCliente")
                    ValorCobrado = 0
                    ValorDoDocumento = Dr1("ValorDoDocumento")
                    Juros = Dr1("Juros")
                    Acrescimos = Dr1("Acrescimos")
                    Descontos = Dr1("Descontos")
                    deducoes = Dr1("Deducoes")
                    If Dr1("Observacoes").ToString.Length > 0 Then
                        Historico = Funcoes.EliminarCaracteresEspeciais(Dr1("Historico")) & ". " & Funcoes.EliminarCaracteresEspeciais(Dr1("Observacoes"))
                    Else
                        Historico = Funcoes.EliminarCaracteresEspeciais(Dr1("Historico"))
                    End If
                    CBancoCliente = Dr1("BancoCliente")
                    CAgenciaCliente = Dr1("AgenciaCliente")
                    CDigitoAgenciaCliente = Dr1("DigitoAgenciaCliente")
                    CCcontaCliente = Dr1("ContaCliente")
                    CDigitoContaCliente = Dr1("DigitoContaCliente")
                    ValorCobrado = ValorDoDocumento + Juros + Acrescimos - Descontos - deducoes
                    TipoPagto = Dr1("TipoPagto")
                    FormaDePagamento = Dr1("Descricao")
                    ''TnumeroDoCheque = Dr1("NumeroDoCheque")
                    Tdescontos = Dr1("Descontos")
                    Tdeducoes = Dr1("Deducoes")
                    TJuros = Dr1("Juros")
                    TAcrescimos = Dr1("Acrescimos")
                    TvalorDoDocumento = Dr1("ValorDodocumento")
                    Tdigito = Dr1("DigitoContaCliente")
                Next

                '' Rotina para leitura de registro buscando a data da baixa (fim).
                '' Dados da empresa

                Dim dr As DataRow
                '' Dados da empresa - fim 
                '' Consultado empresa 
                Sql = " SELECT Cli.Cliente_Id ," & vbCrLf & _
                      "        Cli.Nome, Cli.Cidade," & vbCrLf & _
                      "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                      "        Cli.Endereco , Cli.Cep," & vbCrLf & _
                      "        Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                      "        Cli.Bairro, Cli.Complemento," & vbCrLf & _
                      "        Cli.Numero " & vbCrLf & _
                      "   FROM Clientes Cli " & vbCrLf & _
                      "  WHERE Cli.Cliente_Id = '" & Empresa & "'" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Cli")
                If DS.Tables(0).Rows.Count > 0 Then
                    For Each dr In DS.Tables(0).Rows
                        ENome = dr("Nome")
                        EEndereco = dr("Endereco")
                        ECep = dr("Cep")
                        ECidade = dr("Cidade")
                        EEstado = dr("Estado")
                        ECnpj = dr("Cliente_id")
                        EInscricao = dr("Inscricao")
                        Efone = dr("Telefone")
                        EBairro = dr("Bairro")
                        EComplemento = dr("Complemento")
                        ENumero = dr("Numero")
                        Exit For
                    Next
                End If
                '' Consultando Empresa - fim 
                ''**************************************************************

                '' Dados do Cliente - fim 
                '' Consultado Cliente 
                Sql = " SELECT Cli.Cliente_Id ," & vbCrLf & _
                      "        Cli.Nome, Cli.Cidade," & vbCrLf & _
                      "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                      "        Cli.Endereco , Cli.Cep," & vbCrLf & _
                      "        Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                      "        Cli.Bairro, Cli.Complemento," & vbCrLf & _
                      "        Cli.Numero " & vbCrLf & _
                      "   FROM Clientes Cli " & vbCrLf & _
                      "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf & _
                      "    AND Cli.Endereco_id = " & EndCliente & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Clientes")
                If DS.Tables(0).Rows.Count > 0 Then
                    For Each dr In DS.Tables(0).Rows
                        CNome = dr("Nome")
                        CEndereco = dr("Endereco")
                        CCep = dr("Cep")
                        CCidade = dr("Cidade")
                        CEstado = dr("Estado")
                        CCnpj = dr("Cliente_id")
                        CInscricao = dr("Inscricao")
                        Cfone = dr("Telefone")
                        CBairro = dr("Bairro")
                        CComplemento = dr("Complemento")
                        CNumero = dr("Numero")
                        Exit For
                    Next
                End If
                ' Cria Data Sete que vai ser utilizado no relatorio
                ' Move campos para o Data Set.
                ' Move campos da Empresa
                row = dtRecibo.NewRow()
                row("ENome") = ENome
                row("EEndereco") = EEndereco
                row("ECep") = ECep
                row("ECidade") = ECidade
                row("EEstado") = EEstado
                row("ECnpj") = ECnpj
                row("EInscricao") = EInscricao
                row("EFone") = Efone
                row("EBairro") = EBairro
                row("EComplemento") = EComplemento
                row("ENumero") = ENumero
                '' Move campos do Cliente / fornecedor
                row("CNome") = CNome
                row("CEndereco") = CEndereco
                row("CCep") = CCep
                row("CCidade") = CCidade
                row("CEstado") = CEstado
                row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
                row("CInscricao") = CInscricao
                row("CFone") = Cfone
                row("CBairro") = CBairro
                row("CComplemento") = CComplemento
                row("CNumero") = CNumero
                '' Move campos do Titulo
                row("Tnumtit") = Registro
                row("TValor") = ValorCobrado
                Dim valcobradostr As String
                valcobradostr = CStr(ValorCobrado)
                txtValorCobrado.Text = ValorCobrado

                Valor = Replace(txtValorCobrado.Text, ".", "")

                ''* Rotina de extenso inicio
                yextenso = "("
                yextenso &= UCase(Funcoes.Extenso(txtValorCobrado.Text(), "Real", "Reais"))
                yextenso &= " *"
                xextenso = yextenso
                For j = 1 To (120 - Len(xextenso))
                    xextenso &= " *"
                Next
                xextenso &= ")"
                row("TExtenso") = xextenso
                ''* Rotina de extenso fim 
                row("THistorico") = Historico
                row("TDia") = Day(DataBaixa)
                row("TAno") = Year(DataBaixa)

                row("TMes") = Month(DataBaixa)
                ''* Rotina do Mes inicio
                mes = Month(DataBaixa)
                If mes = 1 Then
                    MESX = "JANEIRO"
                End If

                If mes = 2 Then
                    MESX = "FEVEREIRO"
                End If

                If mes = 3 Then
                    MESX = "MARCO"
                End If

                If mes = 4 Then
                    MESX = "ABRIL"
                End If

                If mes = 5 Then
                    MESX = "MAIO"
                End If

                If mes = 6 Then
                    MESX = "JUNHO"
                End If

                If mes = 7 Then
                    MESX = "JULHO"
                End If

                If mes = 8 Then
                    MESX = "AGOSTO"
                End If

                If mes = 9 Then
                    MESX = "SETEMBRO"
                End If

                If mes = 10 Then
                    MESX = "OUTUBRO"
                End If

                If mes = 11 Then
                    MESX = "NOVEMBRO"
                End If

                If mes = 12 Then
                    MESX = "DEZEMBRO"
                End If

                row("TMes") = MESX
                ''* rotina do mes fim 


                row("TFormaPagto") = FormaDePagamento
                row("TBanco") = CBancoCliente
                row("TAgencia") = CAgenciaCliente
                row("TConta") = CCcontaCliente
                row("TVencimento") = dataVencimento
                row("TBaixa") = DataBaixa

                '' calculo de registro . 
                RegistroI = Registro
                RegistroS = " "
                RegistroS = Trim(RegistroS)
                If Len(RegistroS) < 6 Then
                    For k = 1 To (6 - Len(Trim(RegistroI)))
                        RegistroS &= "0"
                    Next
                End If
                row("Trecibo") = Trim(RegistroS) & (Trim(RegistroI))
                row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))
                row("TValorDoDocumento") = TvalorDoDocumento
                row("TDescontos") = Tdescontos
                row("TDeducoes") = Tdeducoes
                row("TJuros") = TJuros
                row("TAcrescimos") = TAcrescimos
                row("TDigito") = Tdigito
                row("TDigitoAgencia") = CDigitoAgenciaCliente
                dtRecibo.Rows.Add(row)

                ''dsRecibo.Tables.Add(dtRecibo)
                '' Move campos para o Data Fim
                ''
                ''Feito ate aqui falta incluir relatorio e acertar forma de pagamento . 
                '' e corrigir msg de erros
            End If
            i = i + 1
        End While
        If i > 0 Then
            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("xNome", ENome)

            Funcoes.BindReport(Me.Page, dsRecibo, "Cr_ReciboReceber", eExportType.PDF, parameters)
        End If



        ''Imagem
        'Dim dtImagem As New DataTable("Images")
        'dtImagem.Columns.Add("path", GetType(String))
        'dtImagem.Columns.Add("image", GetType(System.Byte()))

        'Dim drImagem As DataRow = dtImagem.NewRow()
        'Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & Session("ssImagemEmpresa"))

        'drImagem("path") = strCaminhoImagem
        'drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        'dtImagem.Rows.Add(drImagem)

        'dsRecibo.Tables.Add(dtImagem)

        'Dim crpt As New ReportDocument()

        'Try
        '    'Dim crpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        '    crpt.FileName = Server.MapPath("~/Reports/Cr_ReciboReceber.rpt")
        '    crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        '    Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        '    Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        '    Dim arquivo As String = NomeArquivo

        '    crpt.SetDataSource(dsRecibo)

        '    Dim crparametervalues As ParameterValues
        '    Dim crparameterdiscretevalue As ParameterDiscreteValue
        '    Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
        '    Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

        '    crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

        '    crparameterfielddefinition = crparameterfielddefinitions.Item("XNome")
        '    crparametervalues = crparameterfielddefinition.CurrentValues
        '    crparameterdiscretevalue = New ParameterDiscreteValue
        '    crparameterdiscretevalue.Value = ENome
        '    crparametervalues.Add(crparameterdiscretevalue)
        '    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

        '    If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        '    crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
        '    Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)

        'Catch ex As Exception
        '    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        'Finally
        '    crpt.Close()
        '    crpt.Dispose()
        'End Try
    End Sub

    Protected Sub cmdAdiantamento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAdiantamento.Click
        If DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
        ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Cliente não foi selecionado!")
        ElseIf String.IsNullOrWhiteSpace(ddlMoeda.SelectedValue) Then
            MsgBox(Me.Page, "Selecione uma Moeda")
        ElseIf Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            Dim Parametros As New Hashtable
            Parametros.Add("Titulo", IIf(Not IsNumeric(txtRegistro.Text), 0, txtRegistro.Text))

            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            Parametros.Add("Empresa", campo(0))
            Parametros.Add("EndEmpresa", campo(1))

            Cliente = txtCodigoCliente.Value
            campo = Cliente.Split("-")
            Parametros.Add("Cliente", campo(0))
            Parametros.Add("EndCliente", campo(1))

            Parametros.Add("Pedido", txtPedido.Text)
            Parametros.Add("PedidoCliente", txtPedido.Text + " - " + txtCliente.Text)

            Dim cart As New CarteiraFinanceira(ddlCarteiras.SelectedValue)
            Parametros.Add("ContaContabil", cart.CodigoContaCliente)
            Parametros.Add("ContaContabilDescricao", cart.CodigoContaCliente + " - " + cart.Descricao)


            Parametros.Add("Moeda", ddlMoeda.SelectedValue)
            Parametros.Add("DescMoeda", ddlMoeda.SelectedItem.Text)
            Parametros.Add("Formulario", "Financeiro")

            Session("Parametros" & HID.Value) = Parametros
            ucConsultaAdiantamentos.BindGridView()
            Popup.ConsultaDeAdiantamentos(Me.Page, "objContasAReceberAdto" & HID.Value)
        End If
    End Sub

    Private Sub ExcluirTitulo()
        If Funcoes.VerificaPermissao("ContasAReceber", "EXCLUIR") Then
            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            If Funcoes.VerificaAcesso(campo(0), campo(1), txtMovimento.Text, "Financeiro") = False Then
                MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para a empresa" & campo(0))
            Else
                Dim SqlArray As New ArrayList

                Dim dsNFxTitulo = CarregarNotaFiscalXTitulo()

                If txtRegistro.Text = "" Then
                    MsgBox(Me.Page, "Informe o número do Registro para Excluir...")
                ElseIf txtPedido.Text <> "" And txtPedido.Text > 0 Then
                    MsgBox(Me.Page, "Título vinculado a Pedido não pode ser excluído...")
                ElseIf Not dsNFxTitulo.Tables Is Nothing AndAlso dsNFxTitulo.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Registro não pode ser Excluído, pois existe uma Nota " & dsNFxTitulo.Tables(0).Rows(0)("Nota_Id") & " vinculada a este.")
                Else
                    Registro = txtRegistro.Text

                    Sql = " UPDATE ContasAReceber" & vbCrLf & _
                          "    SET Situacao = 3, " & vbCrLf & _
                          "        UsuarioCancelamento = '" & Session("ssNomeUsuario") & "'," & vbCrLf & _
                          "        UsuarioCancelamentoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf

                    If txtLiberarTitulo.Value = "S" Then
                        Sql &= ", UsuarioLiberacaoBloqueio = '" & Session("ssNomeUsuario") & "'" 'Usuario que Liberou Titulo
                        Sql &= ", UsuarioLiberacaoBloqueioDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"       'Data da Liberação do Titulo
                    End If

                    If txtLiberarPedido.Value = "S" Then
                        Sql &= ", UsuarioLiberacaoPedido = '" & Session("ssNomeUsuario") & "'" & vbCrLf & _
                               ", UsuarioLiberacaoPedidoDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"         'Data da Liberação do Pedido
                    End If

                    Sql &= " WHERE Registro_ID = " & CStr(Registro)

                    SqlArray.Add(Sql)

                    Sql = "DELETE Razao" & vbCrLf & _
                          " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                    SqlArray.Add(Sql)

                    Sql = "DELETE Adiantamentos " & vbCrLf & _
                          " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                    SqlArray.Add(Sql)

                    Sql = "DELETE FROM AdiantamentosXBaixas " & vbCrLf & _
                          " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                    SqlArray.Add(Sql)

                    Sql = "DELETE TitulosXDesdobrarFornecedor " & vbCrLf & _
                          " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

                    SqlArray.Add(Sql)

                    '-----Consulta Registros Filhos------
                    Sql = " SELECT Registro_id " & vbCrLf & _
                          "   FROM contasAreceber " & vbCrLf & _
                          "  WHERE RegistroMestre = " & txtRegistro.Text

                    Dim dsMestre As New DataSet
                    dsMestre = Banco.ConsultaDataSet(Sql, "Registros")

                    If Not dsMestre Is Nothing AndAlso dsMestre.Tables(0).Rows.Count > 0 Then
                        For Each drFilho As DataRow In dsMestre.Tables(0).Rows
                            Sql = " UPDATE ContasAReceber" & vbCrLf & _
                                  "    SET Situacao = 1" & vbCrLf & _
                                  "        ,Provisao = 2" & vbCrLf & _
                                  "        ,Grupado = 'N'" & vbCrLf & _
                                  "        ,RegistroMestre = 0" & vbCrLf & _
                                  "        ,UsuarioCancelamento = '" & Session("ssNomeUsuario") & "'" & vbCrLf & _
                                  "        ,UsuarioCancelamentoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf

                            If txtLiberarTitulo.Value = "S" Then
                                Sql &= "       ,UsuarioLiberacaoBloqueio = '" & Session("ssNomeUsuario") & "'" & vbCrLf & _
                                       "       ,UsuarioLiberacaoBloqueioDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf       'Data da Liberação do Titulo
                            End If

                            If txtLiberarPedido.Value = "S" Then
                                Sql &= "       ,UsuarioLiberacaoPedido = '" & Session("ssNomeUsuario") & "'" & vbCrLf & _
                                       "       ,UsuarioLiberacaoPedidoDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf          'Data da Liberação do Pedido
                            End If

                            Sql &= "       WHERE Registro_Id = " & drFilho("Registro_id") & vbCrLf

                            SqlArray.Add(Sql)

                            Sql = "DELETE FROM Razao " & vbCrLf & _
                                  " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                            SqlArray.Add(Sql)

                            Sql = "DELETE FROM Adiantamentos" & vbCrLf & _
                                  " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                            SqlArray.Add(Sql)

                            Sql = "DELETE FROM AdiantamentosXBaixas" & vbCrLf & _
                                  " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                            SqlArray.Add(Sql)
                        Next
                    End If

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar(True)
                        MsgBox(Me.Page, "Registro < " & Registro & " > excluído com sucesso.", eTitulo.Sucess)
                    End If
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
        End If
    End Sub

    Protected Sub cmdPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPedido.Click
        If DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
        ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cliente não foi selecionado!")
        ElseIf (txtValorDoDocumento.Text.Length = 0 AndAlso txtValorEmMoeda.Text.Length = 0) OrElse txtValorCobrado.Text.Length = 0 Then
            MsgBox(Me.Page, "Valor do documento em R$ ou U$ e valor líquido para pgto são obrigatórios!")
        Else
            Session("ssCampo" & HID.Value) = "Pedidos"
            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue
            parameters("empresa") = campo(0)
            parameters("enderecoEmpresa") = campo(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedidoCTAREC" & HID.Value)

            ucConsultaPedidos.BindGridView(parameters)
        End If
    End Sub

    Protected Sub imgBloqueio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        If Not Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
            MsgBox(Me.Page, "Usuário sem permissão para liberar Registro")
        ElseIf Not Funcoes.VerificaAcesso(campo(0), campo(1), txtMovimento.Text, "FINANCEIRO") Then
            MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para empresa " & campo(0))
        Else
            Dim Mestre() As String = txtMestre.Text.Split(":")

            If CInt(Trim(Mestre(1))) > 0 Then
                MsgBox(Me.Page, "Registro de Agrupamento não pode ser desbloqueado, para alteração deve desfazer o Agrupamento")
                Exit Sub
            End If

            txtMovimento.Enabled = True
            ShowCalendar(Me.Page, txtMovimento)
            txtProrrogacao.Enabled = True
            ShowCalendar(Me.Page, txtProrrogacao)
            txtLiberarTitulo.Value = "S"
            DdlProvisoes.Enabled = True
            ImgCalcular.Enabled = True
            lnkNovo.Parent.Visible = True
            ddlIndexador.Enabled = True
            txtValorDoDocumento.Enabled = True
            txtDescontos.Enabled = True
            txtDeducoes.Enabled = True
            txtJuros.Enabled = True
            txtAcrescimos.Enabled = True
            txtValorEmMoeda.Enabled = True
            txtDescontosMoeda.Enabled = True
            txtDeducoesMoeda.Enabled = True
            txtJurosMoeda.Enabled = True
            txtAcrescimosMoeda.Enabled = True
        End If
    End Sub

    Protected Sub imgLimparPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
            txtLiberarPedido.Value = "S"
            txtPedido.Text = "0"
            cmdPedido.Enabled = True
        Else
            MsgBox(Me.Page, "Usuario sem permissao para remover o Pedido.")
        End If
    End Sub

    Protected Sub imgLimparAdto_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparAdto.Click
        If Funcoes.VerificaPermissao("ContasAReceber", "LIBERAR") Then
            Dim cart As New CarteiraFinanceira(ddlCarteiras.SelectedValue)

            If Not (cart.isAdiantamento And Not cart.BaixaAdiantamento) Then
                txtNumeroAdto.Text = "0"
                HDSaldoAdiantamento.Value = 0
                cmdAdiantamento.Enabled = True
            Else
                MsgBox(Me.Page, "Este titulo deu origem a um adiantamento. para apaga-lo, caso nao tenha baixas atreladas a ele mude para previsao e salve.")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para liberar registro.")
        End If
    End Sub

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If txtRegistro.Text.Length = 0 Then
            MsgBox(Me.Page, "Consulte o Registro para visualização do Extrato")
        ElseIf DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa do Registro não encontrada")
        ElseIf String.IsNullOrWhiteSpace(txtCliente.Text) Then
            MsgBox(Me.Page, "Cliente do Registro não encontrado")
        ElseIf txtPedido.Text.Length = 0 OrElse txtPedido.Text = "0" Then
            MsgBox(Me.Page, "Registro sem Pedido não pode ser visualizado")
        Else
            Extrato.Emitir(Me.Page, FinanceiroNovo, DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), "T", _
                           txtPedido.Text.Trim, "", 0, "", Nothing, "", False, False, False, False, True)

            'Dim strQueryString As String = "?fim=" & DateTime.Now.ToString("dd/MM/yyyy")
            'strQueryString &= "&empresa=" & DdlEmpresaCliente.SelectedValue
            'strQueryString &= "&cliente=" & Replace(txtCodigoCliente.Value, ";", "-")
            'strQueryString &= "&pedido=" & txtPedido.Text
            'strQueryString &= "&es=ES"

            'campo = DdlEmpresaCliente.SelectedValue.ToString.Split("-")
            'Dim objEmpresa As New [Lib].Negocio.Cliente(campo(0), campo(1))

            'For Each row As [Lib].Negocio.ClientexTipo In objEmpresa.Tipos
            '    If row.CodigoTipo = eTipoCliente.Revenda Then
            '        strQueryString &= "&desprd=S"
            '    End If
            'Next
            'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, "CarregarHTML", "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
        End If
    End Sub

    Protected Sub DdlProvisoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If DdlProvisoes.SelectedIndex = 0 Then Exit Sub

        If Not DdlProvisoes.SelectedValue = 1 AndAlso txtProrrogacao.Enabled = False Then
            txtProrrogacao.Enabled = True
            ShowCalendar(Me.Page, txtProrrogacao)
        End If

        If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
            If (Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value)) Then
                Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                If (strCliente IsNot Nothing AndAlso strCliente.Length > 0) Then
                    Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
                    If objCliente.DesdobrarFornecedor = True Then
                        ucDestinoContabil.Limpar()
                        Dim parameters As New Dictionary(Of String, Object)
                        parameters.Add("tipo", "R")
                        Popup.ConsultaDeDestinoContabil(Me.Page, "objDestinoContabil" & HID.Value)
                        ucDestinoContabil.Carregar(parameters)
                    End If
                End If
            Else
                DdlProvisoes.SelectedIndex = 0
                MsgBox(Me.Page, "É necessário selecionar o campo fornecedor!")
            End If
        End If
    End Sub

    Protected Sub ChkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTitulo As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkTitulo.NamingContainer, GridViewRow)

        If Not chkTitulo.Checked Then
            Dim lstSelecionados As List(Of String) = GridConsultaTitulos.GetSelectedValues("chkSelecionado")
            If lstSelecionados Is Nothing OrElse lstSelecionados.Count <= 0 Then
                HiddenIndexador.Value = String.Empty
            End If
        End If

        If chkTitulo.Checked Then
            Dim strMoeda() As String = GridConsultaTitulos.Rows(row.RowIndex).Cells(8).Text.ToString.Split("-")
            If txtRealDolar.Value.ToString.Length = 0 Then
                txtRealDolar.Value = strMoeda(1)
            End If

            If HiddenIndexador.Value.ToString.Length = 0 Then
                HiddenIndexador.Value = GridConsultaTitulos.Rows(row.RowIndex).Cells(9).Text
            End If

            If Not strMoeda(1) = txtRealDolar.Value Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Agrupamento permitido apenas para Títulos com a mesma moeda.")
            End If

            If Not GridConsultaTitulos.Rows(row.RowIndex).Cells(9).Text = HiddenIndexador.Value Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Agrupamento permitido apenas para Títulos com o mesmo indexador.")
            End If
        End If
        TotalizadorTitulosAgrupados()
    End Sub

    Protected Sub txtProrrogacao_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtProrrogacao.TextChanged
        Try
            MostrarCotacao()

            If (Not String.IsNullOrWhiteSpace(txtRegistro.Text)) AndAlso (Not String.IsNullOrWhiteSpace(ddlMoeda.SelectedValue)) Then
                If ddlIndexador.SelectedValue = 99 Then
                    ValidaValores(True)
                Else
                    Dim vlrMoeda As Decimal = Decimal.Zero
                    Dim sql As String = "SELECT * FROM Cotacoes WHERE Data_Id = '" & CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd") & "' AND Indexador_Id = " & ddlIndexador.SelectedValue & " "
                    Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Cotacoes")
                    If (ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0) Then
                        Dim row As DataRow = ds.Tables(0).Rows(0)
                        vlrMoeda = CDec(row("Indice"))
                    End If

                    If ddlMoeda.SelectedValue = 3 Then
                        txtValorCobrado.Text = String.Format("{0:N2}", (CDec(txtValorEmMoeda.Text) * vlrMoeda))
                        Dim vlrDiff As Decimal = CDec(txtValorCobrado.Text) - CDec(txtValorDoDocumento.Text)
                        If (vlrDiff > 0) Then
                            txtAcrescimos.Text = String.Format("{0:N2}", vlrDiff)
                        Else
                            txtDeducoes.Text = String.Format("{0:N2}", (vlrDiff * -1))
                        End If
                    Else
                        txtValorCobradoMoeda.Text = String.Format("{0:N2}", (CDec(txtValorDoDocumento.Text) / vlrMoeda))
                        Dim vlrDiff As Decimal = CDec(txtValorCobradoMoeda.Text) - CDec(txtValorEmMoeda.Text)
                        If (vlrDiff > 0) Then
                            txtAcrescimosMoeda.Text = String.Format("{0:N2}", vlrDiff)
                        Else
                            txtDeducoesMoeda.Text = String.Format("{0:N2}", (vlrDiff * -1))
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("ContasAReceber", "GRAVAR") Then
            Try
                GravaTitulo()
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Else
            MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If txtRegistro.Text <> "" And txtRegistro.Text <> "0" Then
            If DdlProvisoes.SelectedValue = 1 Then
                MsgBox(Me.Page, "Título baixado não pode ser excluído")
            Else
                ExcluirTitulo()
            End If
        End If
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("ContasAReceber", "LEITURA") Then
            Dim registro As String = txtRegistro.Text
            Limpar(True)
            txtRegistro.Text = registro
            ConsultaContasAReceber()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        chkManterLancamento.Checked = False
        Limpar(True)
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        EmitirRecibo()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ContasAReceber")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkReprogramar_Click(sender As Object, e As EventArgs) Handles lnkReprogramar.Click
        pnlReprogramaVencimentos.Visible = True
        txtNovoVencimento.Focus()
    End Sub

    Protected Sub btnNovoVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNovoVencimento.Click
        Try
            If String.IsNullOrWhiteSpace(txtNovoVencimento.Text) Then
                MsgBox(Me.Page, "Informe o novo vencimento para os Títulos")
            ElseIf Not IsDate(txtNovoVencimento.Text) Then
                MsgBox(Me.Page, "Data para novo vencimento não é válida")
            Else
                Dim sql As String = String.Empty
                Dim ischeckedExist As Boolean = False
                Dim lstTitulosSelecionados As New List(Of String)

                For Each row As GridViewRow In GridConsultaTitulos.Rows
                    If CType(row.FindControl("chkReprogramar"), CheckBox).Checked Then
                        ischeckedExist = True
                        lstTitulosSelecionados.Add(row.Cells(3).Text())
                    End If
                Next

                sql = "UPDATE ContasAReceber " & vbCrLf & _
                      "   Set Prorrogacao          = '" & CDate(txtNovoVencimento.Text).ToString("yyyy/MM/dd") & "' " & vbCrLf & _
                      "     , UsuarioAlteracao     = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                      "     , UsuarioAlteracaoData = GETDATE()" & vbCrLf & _
                      " Where Registro_Id in (" & String.Join(", ", lstTitulosSelecionados) & ")"

                If Banco.GravaBanco(sql) Then
                    LimparConsultaTitulos(True)
                    MsgBox(Me.Page, "Titulo(s): " & String.Join(", ", lstTitulosSelecionados) & " Reprogramados com Sucesso.", eTitulo.Sucess)

                Else
                    MsgBox(Me.Page, "Não foi(ram) selecionado(s) titulo(s) para reprogramação de vencimento")

                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCancelarNonoVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancelarNonoVencimento.Click
        Try
            LimparConsultaTitulos(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class