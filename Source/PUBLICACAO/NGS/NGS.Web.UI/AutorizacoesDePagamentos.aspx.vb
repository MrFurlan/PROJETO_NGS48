Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AutorizacoesDePagamentos
    Inherits BasePage



#Region "Atributos/ Propriedades"
    Dim Sql As String
    Dim SqlArray As New ArrayList
#End Region

#Region "Eventos"

#Region "Page_Load"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                LimparCampos("T")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

#End Region

#Region "LinkButton"
    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "LEITURA") Then
            Consultar_Registros()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro(s).")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        LimparCampos("T")
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "RELATORIO") Then
                Dim DS As New DataSet
                Dim DSEmp As New DataSet
                Dim Campo As String()

                Dim UnidadeDeNegocio As String = String.Empty
                Dim NomeUnidadeDeNegocio As String = "Consolidado"

                Dim Empresa As String = String.Empty
                Dim EndEmpresa As String = String.Empty

                Dim NomeEmpresa As String = String.Empty
                Dim CidadeEmpresa As String = String.Empty
                Dim EstadoEmpresa As String = String.Empty
                'Dim GrupoEmpresa As String = String.Empty 

                Dim Parametros As String = ""

                Dim Cliente As String = String.Empty
                Dim EndCliente As String = String.Empty

                If ddlUnidadeDeNegocio.Text <> "" Then
                    UnidadeDeNegocio = ddlUnidadeDeNegocio.SelectedValue
                Else
                    UnidadeDeNegocio = ""                   'UnidadeDeNegocio
                End If

                If ddlEmpresa.Text <> "" Then
                    Cliente = ddlEmpresa.SelectedValue
                    Campo = Cliente.Split("-")
                    Empresa = Campo(0)                      'EmpresaCliente
                    EndEmpresa = Campo(1)                   'Endereco Empresa Cliente
                    Parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                Else
                    Empresa = ""                            'Empresa Cliente
                    EndEmpresa = 0                          'Endereco Empresa Cliente
                End If

                If txtCodigoCliente.Value.Length > 0 Then
                    Campo = txtCodigoCliente.Value.ToString.Split("-")
                    Cliente = Campo(0)                      'Cliente
                    EndCliente = Campo(1)                   'Endereco Cliente
                    Parametros &= "Cliente: " & txtCliente.Text & vbCrLf
                Else
                    Cliente = ""                            'Cliente
                    EndCliente = 0                          'Endereco Cliente
                End If

                Dim DataInicial As String = CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd")
                Dim DataFinal As String = CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd")

                Dim DataInicialD As String = CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("dd/MM/yyyy")
                Dim DataFinalD As String = CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("dd/MM/yyyy")

                '--Empresa---------------------------------------------------

                Sql = " SELECT Clientes.Cliente_Id AS Cliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
                      "   FROM Clientes " & vbCrLf & _
                      "  WHERE Clientes.Cliente_Id + CONVERT(NVARCHAR,Clientes.Endereco_Id) = '" & Empresa & EndEmpresa & "' " & vbCrLf

                DSEmp = Banco.ConsultaDataSet(Sql, "Clientes")
                If DSEmp.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In DSEmp.Tables(0).Rows
                        NomeEmpresa = dr("Nome")
                        CidadeEmpresa = dr("Cidade")
                        EstadoEmpresa = dr("Estado")
                        Exit For
                    Next
                Else
                    NomeEmpresa = HttpContext.Current.Session("ssNomeEmpresa")
                    CidadeEmpresa = "Geral"
                    EstadoEmpresa = ""
                End If

                DS = SqlConsulta()

                Dim param As New Dictionary(Of String, Object)
                param.Add("Nome", NomeEmpresa)
                param.Add("Cidade", CidadeEmpresa)
                param.Add("Estado", EstadoEmpresa)
                param.Add("DataInicial", DataInicialD)
                param.Add("DataFinal", DataFinalD)

                If Pdf Then
                    Funcoes.BindReport(Me.Page, DS, "Cr_RelAutorizacoesDePgtos", eExportType.PDF, param)
                Else
                    Funcoes.BindReport(Me.Page, DS, "Cr_RelAutorizacoesDePgtos", eExportType.ExcelCrystal, param)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para Relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
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

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedido.Click
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

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedido" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#End Region

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteAUTPGTO" & HID.Value) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteAUTPGTO" & HID.Value)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteAUTPGTO" & HID.Value)
        ElseIf Session("objPedido" & HID.Value) IsNot Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = p.Codigo
            Session.Remove("objPedido" & HID.Value)
        End If
    End Sub

    Private Sub LimparCampos(ByVal Tipo As String)
        If Tipo = "T" Then
            txtPeriodoInicialConsultaTitulos.Text = Today.ToString("dd/MM/yyyy")
            txtPeriodoFinalConsultaTitulos.Text = Today.ToString("dd/MM/yyyy")
            txtCliente.Text = ""
            txtPedido.Text = String.Empty
            txtTitulo.Text = String.Empty
            txtCodigoCliente.Value = ""
            Session.Remove("objClienteAUTPGTO" & HID.Value)
            lblTotalRegistroReais.Text = ""
            lblTotalRegistroDolar.Text = ""
            chkConsolidarEmpresa.Checked = False
            LinkLiberar.Parent.Visible = False
        End If

        gridRegistro.Parent.Visible = False
        gridProduto.Parent.Visible = False
        gridProdutoXRegistros.Parent.Visible = False
        chkProdutoXRegistrosXTodos.Enabled = False
        imgBuscarCliente.Enabled = False

        txtValor.Text = "0,00"
        lblTotalProdutosXRegistros.Text = "0,00"
        Session.Remove("objAUTPGTO" & Session.SessionID)

        gridRegistro.DataSource = Nothing
        gridRegistro.DataBind()

        gridProduto.DataSource = Nothing
        gridProduto.DataBind()

        gridProdutoXRegistros.DataSource() = Nothing
        gridProdutoXRegistros.DataBind()

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresa()
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteAUTPGTO" & HID.Value, "txtNome")
    End Sub

    Private Function SqlConsulta() As DataSet
        Dim ds As New DataSet
        Sql = "SELECT ContasAPagar.Registro_Id AS Registro, (ContasAPagar.Empresa + '-' + left(Emp.Nome,20) + '-' + rtrim(Emp.Cidade) + '/' + Emp.Estado) AS Empresa, convert(varchar(10),ContasAPagar.Prorrogacao,103) as Vencimento, " & vbCrLf & _
              "       Cli.Nome AS Cliente, " & vbCrLf & _
              "       CASE " & vbCrLf & _
              "       	  WHEN ContasAPagar.Moeda = 1 " & vbCrLf & _
              "       		  THEN ROUND(ISNULL(ContasAPagar.ValorLiquido, 0),2) " & vbCrLf & _
              "       	      ELSE 0 " & vbCrLf & _
              "       	  END AS Valor, " & vbCrLf & _
              "       CASE " & vbCrLf & _
              "       	  WHEN ContasAPagar.Moeda <> 1 " & vbCrLf & _
              "       		   THEN ROUND(ISNULL(ContasAPagar.MoedaValorLiquido, 0),2)  " & vbCrLf & _
              "                ELSE 0 " & vbCrLf & _
              "       	   END AS Dolar, " & vbCrLf & _
              "       ContasAPagar.Historico as Historico , isnull(ContasAPagar.UsuarioLiberacao,'') as UsuarioLiberacao, Provisao, Moeda, Indexador, " & vbCrLf & _
              "       isnull(ContasAPagar.Pedido, 0) AS Pedido, " & vbCrLf & _
              "       isnull(dbo.ProdutosPedido(ContasAPagar.EmpresaPedido,ContasAPagar.EndEmpresaPedido,ContasAPagar.Pedido),'') AS NomeProduto, " & vbCrLf & _
              "       convert(numeric(18,2),0) as ValorLiberado, " & vbCrLf & _
              "       convert(numeric(18,2),0) as ValorNaoLiberado, " & vbCrLf & _
              "       convert(numeric(18,2),0) as MoedaLiberado, " & vbCrLf & _
              "       convert(numeric(18,2),0) as MoedaNaoLiberado " & vbCrLf & _
              "  INTO #Temp " & vbCrLf & _
              "  FROM ContasAPagar " & vbCrLf & _
              " INNER JOIN Clientes AS Cli " & vbCrLf & _
              "    ON ContasAPagar.Cliente    = Cli.Cliente_Id " & vbCrLf & _
              "   AND ContasAPagar.EndCliente = Cli.Endereco_Id " & vbCrLf & _
              " INNER JOIN Clientes AS Emp " & vbCrLf & _
              "    ON ContasAPagar.Empresa    = Emp.Cliente_Id " & vbCrLf & _
              "   AND ContasAPagar.EndEmpresa = Emp.Endereco_Id " & vbCrLf & _
              "  LEFT JOIN Produtos " & vbCrLf & _
              "    ON ContasAPagar.Carteira = Produtos.Produto_Id " & vbCrLf & _
              " WHERE ContasApagar.Situacao = 1 " & vbCrLf & _
              "   AND ContasApagar.Provisao = 2 " & vbCrLf & _
              "   AND ISNULL(ContasApagar.RegistroMestre,0) = 0 " & vbCrLf

        If ddlUnidadeDeNegocio.SelectedIndex > 0 Then
            Sql &= "   AND ContasAPagar.UnidadeDeNegocio = '" & ddlUnidadeDeNegocio.SelectedValue & "' " & vbCrLf
        End If

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

            If chkConsolidarEmpresa.Checked Then
                Sql &= "   AND LEFT(ContasAPagar.Empresa, 8) = '" & Left(Empresa(0), 8) & "'" & vbCrLf
            Else
                Sql &= "   AND ContasAPagar.Empresa = '" & Empresa(0) & "'" & vbCrLf & _
                       "   AND ContasAPagar.EndEmpresa = " & Empresa(1) & vbCrLf
            End If

        End If

        If txtCliente.Text.Length > 0 Then
            Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= "   AND ContasAPagar.Cliente = '" & Cliente(0) & "'" & vbCrLf & _
                   "   AND ContasAPagar.EndCliente = " & Cliente(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso IsNumeric(txtPedido.Text) Then
            Sql &= "   AND ContasAPagar.Pedido = " & txtPedido.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtTitulo.Text) AndAlso IsNumeric(txtTitulo.Text) Then
            Sql &= "   AND ContasAPagar.Registro_Id = " & txtTitulo.Text & vbCrLf
        Else
            If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
                Sql &= "   AND ContasAPagar.Prorrogacao between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf
            End If
        End If

        If rdALiberar.Checked Then
            Sql &= "   AND len(isnull(ContasAPagar.UsuarioLiberacao,'')) = 0 " & vbCrLf
        ElseIf rdLiberado.Checked Then
            Sql &= "   AND len(isnull(ContasAPagar.UsuarioLiberacao,'')) > 0 " & vbCrLf
        End If

        Sql &= vbCrLf & _
                "UPDATE #Temp " & vbCrLf & _
                "   SET NomeProduto = case when isnull(NomeProduto,'') = ''	then 'OUTROS' else NomeProduto end, " & vbCrLf & _
                "       ValorLiberado = case when isnull(UsuarioLiberacao,'') <> ''	then Valor else 0 end, " & vbCrLf & _
                "       ValorNaoLiberado = case when isnull(UsuarioLiberacao,'') = '' then  Valor else 0 end, " & vbCrLf & _
                "       MoedaLiberado = case when isnull(UsuarioLiberacao,'') <> '' then Dolar else 0 end, " & vbCrLf & _
                "       MoedaNaoLiberado = case	when isnull(UsuarioLiberacao,'') = '' then Dolar else 0 end " & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "select NomeProduto, " & vbCrLf & _
                "       sum(Valor) as Valor, " & vbCrLf & _
                "       sum(Dolar) as Dolar, " & vbCrLf & _
                "       sum(ValorLiberado) as ValorLiberado, " & vbCrLf & _
                "       sum(ValorNaoLiberado) as ValorNaoLiberado, " & vbCrLf & _
                "       sum(MoedaLiberado) as MoedaLiberado, " & vbCrLf & _
                "       sum(MoedaNaoLiberado) as MoedaNaoLiberado " & vbCrLf & _
                "  Into #TotalProduto " & vbCrLf & _
                "  from #Temp " & vbCrLf & _
                " group by NomeProduto " & vbCrLf & _
                "" & vbCrLf & _
                "select Registro, Empresa, Vencimento, Cliente, Dolar, Valor, Historico, " & vbCrLf & _
                "       UsuarioLiberacao, Provisao, Moeda, Indexador, Pedido, NomeProduto " & vbCrLf & _
                "  from #Temp " & vbCrLf & _
                " ORDER BY Vencimento, Registro " & vbCrLf & _
                "" & vbCrLf & _
                "select NomeProduto, Valor, Dolar, ValorLiberado, ValorNaoLiberado, MoedaLiberado, MoedaNaoLiberado " & vbCrLf & _
                "  from #TotalProduto " & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "AutPgtos")

        Return ds
    End Function

    Private Function SqlConsultaFinanceiroNovo() As DataSet
        Dim ds As New DataSet
        Sql = " SELECT " & vbCrLf & _
            " T.Titulo_Id AS Registro, " & vbCrLf & _
            " T.Empresa + ' - ' + LEFT(EMP.Nome, 20) + ' - ' + RTRIM(Emp.Cidade) + ' - ' + Emp.Estado AS Empresa, " & vbCrLf & _
            " CONVERT(VARCHAR(10),T.Reprogramacao,103) as Vencimento, " & vbCrLf & _
            " Cli.Nome AS Cliente, " & vbCrLf & _
            " CASE WHEN T.RecPag = 'P' THEN " & vbCrLf & _
            " (SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C') " & vbCrLf & _
            " END AS Dolar, " & vbCrLf & _
            " CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf & _
            " (SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C')  " & vbCrLf & _
            " END AS Valor, " & vbCrLf & _
            " T.Historico, " & vbCrLf & _
            " ISNULL(TH.Usuario, '') AS UsuarioLiberacao, " & vbCrLf & _
            " T.Provisao, " & vbCrLf & _
            " T.Moeda, " & vbCrLf & _
            " T.Indexador, " & vbCrLf & _
            " T.Pedido, " & vbCrLf & _
            " ISNULL(dbo.ProdutosPedido(T.Empresa,T.EndEmpresa,T.Pedido),'') AS NomeProduto, " & vbCrLf & _
            " CONVERT(NUMERIC(18,2),0) AS ValorLiberado, " & vbCrLf & _
            " CONVERT(NUMERIC(18,2),0) AS ValorNaoLiberado, " & vbCrLf & _
            " CONVERT(NUMERIC(18,2),0) AS MoedaLiberado, " & vbCrLf & _
            " CONVERT(NUMERIC(18,2),0) AS MoedaNaoLiberado " & vbCrLf & _
            " INTO #Temp " & vbCrLf & _
            " FROM Titulos AS T " & vbCrLf & _
            " INNER JOIN Clientes AS Emp " & vbCrLf & _
            " ON T.Empresa = Emp.Cliente_Id " & vbCrLf & _
            " AND T.EndEmpresa = Emp.Endereco_Id " & vbCrLf & _
            " INNER JOIN Clientes AS Cli " & vbCrLf & _
            " ON T.Empresa = Cli.Cliente_Id " & vbCrLf & _
            " AND T.EndEmpresa = Cli.Endereco_Id " & vbCrLf & _
            " LEFT JOIN TitulosXHistorico AS TH " & vbCrLf & _
            " ON T.Titulo_Id = TH.Titulo_Id " & vbCrLf & _
            " WHERE T.RecPag = 'P' " & vbCrLf & _
            " AND T.Situacao = 1 " & vbCrLf & _
            " AND T.Provisao = 3 " & vbCrLf & _
            " AND T.RegistroMestre IS NULL " & vbCrLf

        If ddlUnidadeDeNegocio.SelectedIndex > 0 Then
            Sql &= "   AND T.UnidadeDeNegocio = '" & ddlUnidadeDeNegocio.SelectedValue & "' " & vbCrLf
        End If

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

            If chkConsolidarEmpresa.Checked Then
                Sql &= "   and left(T.Empresa, 8) = '" & Left(Empresa(0), 8) & "'" & vbCrLf
            Else
                Sql &= "   and T.Empresa = '" & Empresa(0) & "'" & vbCrLf & _
                       "   and T.EndEmpresa = " & Empresa(1) & vbCrLf
            End If
        End If

        If txtCliente.Text.Length > 0 Then
            Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
            Sql &= "   AND T.CliFor = '" & Cliente(0) & "'" & vbCrLf & _
                   "   AND T.EnderecoCliFor = " & Cliente(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso IsNumeric(txtPedido.Text) Then
            Sql &= "  And T.Pedido = " & txtPedido.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtTitulo.Text) AndAlso IsNumeric(txtTitulo.Text) Then
            Sql &= "  And T.Titulo_Id = " & txtTitulo.Text & vbCrLf
        Else
            If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
                Sql &= "   AND T.Reprogramacao between '" & Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "yyyy-MM-dd") & "'" & vbCrLf
            End If
        End If

        If rdALiberar.Checked Then
            Sql &= "   AND len(isnull(TH.Usuario,'')) = 0 " & vbCrLf
        ElseIf rdLiberado.Checked Then
            Sql &= "   AND len(isnull(TH.Usuario,'')) > 0 " & vbCrLf
        End If

        Sql &= vbCrLf & _
                "update #Temp set" & vbCrLf & _
                "   NomeProduto = case when isnull(NomeProduto,'') = ''	then 'OUTROS' else NomeProduto end, " & vbCrLf & _
                "   ValorLiberado = case when isnull(UsuarioLiberacao,'') <> ''	then Valor else 0 end, " & vbCrLf & _
                "   ValorNaoLiberado = case when isnull(UsuarioLiberacao,'') = '' then  Valor else 0 end, " & vbCrLf & _
                "   MoedaLiberado = case when isnull(UsuarioLiberacao,'') <> '' then Dolar else 0 end, " & vbCrLf & _
                "   MoedaNaoLiberado = case	when isnull(UsuarioLiberacao,'') = '' then Dolar else 0 end " & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "select NomeProduto, " & vbCrLf & _
                "   sum(Valor) as Valor, " & vbCrLf & _
                "   sum(Dolar) as Dolar, " & vbCrLf & _
                "   sum(ValorLiberado) as ValorLiberado, " & vbCrLf & _
                "   sum(ValorNaoLiberado) as ValorNaoLiberado, " & vbCrLf & _
                "   sum(MoedaLiberado) as MoedaLiberado, " & vbCrLf & _
                "   sum(MoedaNaoLiberado) as MoedaNaoLiberado " & vbCrLf & _
                "   Into #TotalProduto " & vbCrLf & _
                "   from #Temp " & vbCrLf & _
                "   group by NomeProduto " & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "select Registro, Empresa, Vencimento, Cliente, Dolar, Valor, Historico, " & vbCrLf & _
                "   UsuarioLiberacao, Provisao, Moeda, Indexador, Pedido, NomeProduto " & vbCrLf & _
                "   from #Temp " & vbCrLf & _
                "   ORDER BY Vencimento, Registro " & vbCrLf & _
                "" & vbCrLf & _
                "" & vbCrLf & _
                "select NomeProduto, Valor, Dolar, ValorLiberado, ValorNaoLiberado, MoedaLiberado, MoedaNaoLiberado " & vbCrLf & _
                "from #TotalProduto " & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "AutPgtos")

        Return ds
    End Function

    Private Sub Consultar_Registros()
        Dim rvlrLiberado As Decimal = 0
        Dim rvlrALiberar As Decimal = 0
        Dim dvlrLiberado As Decimal = 0
        Dim dvlrALiberar As Decimal = 0

        Dim ds As New DataSet
        If FinanceiroNovo Then
            ds = SqlConsultaFinanceiroNovo()
        Else
            ds = SqlConsulta()
        End If

        If ds Is Nothing AndAlso ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum Registro foi encontrado.")
            Exit Sub
        Else
            For Each Dr As DataRow In ds.Tables(0).Rows
                If Dr("Provisao") = "2" AndAlso Dr("Moeda") = 3 Then
                    Dr("Cliente") = "$" & Dr("Cliente")
                End If
            Next

            gridRegistro.Parent.Visible = True
            gridProduto.Parent.Visible = True
            gridProdutoXRegistros.Parent.Visible = True

            gridRegistro.DataSource = ds.Tables(0)
            gridRegistro.DataBind()

            gridProduto.DataSource = ds.Tables(1)
            gridProduto.DataBind()

            For Each dr As DataRow In ds.Tables("AutPgtos").Rows
                If dr("UsuarioLiberacao").ToString.Length = 0 Then
                    rvlrALiberar += dr("Valor")
                    dvlrALiberar += dr("Dolar")
                Else
                    rvlrLiberado += dr("Valor")
                    dvlrLiberado += dr("Dolar")
                End If
            Next
        End If

        Session("objAUTPGTO" & Session.SessionID) = ds

        Dim i As Integer = 0
        While i < gridRegistro.Rows.Count
            Dim campoTeste = gridRegistro.Rows(i).Cells(12).Text()

            If Mid(gridRegistro.Rows(i).Cells(4).Text(), 1, 1) = "$" Then
                gridRegistro.Rows(i).Cells(8).ForeColor = Drawing.Color.Red
                gridRegistro.Rows(i).Cells(9).ForeColor = Drawing.Color.Red
                gridRegistro.Rows(i).Cells(4).Text() = gridRegistro.Rows(i).Cells(4).Text().Replace("$", "")
            End If

            If campoTeste = "&nbsp;" Or campoTeste = "" Or campoTeste = " " Then
                CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = False
            Else
                CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = True
            End If

            CType(gridRegistro.Rows(i).FindControl("imgSituacao"), ImageButton).ImageUrl = "~/images/question.jpg"

            i += 1
        End While

        LinkLiberar.Parent.Visible = True
        lblTotalRegistroReais.Text = "(Reais) À Liberar: " & rvlrALiberar.ToString("N2") & "  -  Liberado: " & rvlrLiberado.ToString("N2") & "  -  Total: " & (rvlrALiberar + rvlrLiberado).ToString("N2")
        lblTotalRegistroDolar.Text = "(Dólar) À Liberar: " & dvlrALiberar.ToString("N2") & "  -  Liberado: " & dvlrLiberado.ToString("N2") & "  -  Total: " & (dvlrALiberar + dvlrLiberado).ToString("N2")
    End Sub

    Protected Sub LinkLiberar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then

            SqlArray.Clear()
            Dim i As Integer = 0
            While i < gridRegistro.Rows.Count
                Sql = ""
                If CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = True Then
                    If String.IsNullOrWhiteSpace(Server.HtmlDecode(gridRegistro.Rows(i).Cells(12).Text)) Then
                        If FinanceiroNovo Then
                            Dim TituloXHistorico As New Novo.TituloXHistorico()
                            TituloXHistorico.Acao = "LIBERAR"
                            TituloXHistorico.CodigoTitulo = gridRegistro.Rows(i).Cells(1).Text
                            TituloXHistorico.Data = DateTime.Now
                            TituloXHistorico.Historico = "LIBERAÇÃO DE TíTULO "
                            TituloXHistorico.Usuario = Session("ssNomeUsuario")
                            TituloXHistorico.Salvar()
                        Else
                            Sql = "UPDATE ContasAPagar " & vbCrLf & _
                              "   SET UsuarioLiberacao = '" & Session("ssNomeUsuario") & "', " & vbCrLf & _
                              "       UsuarioLiberacaoData = '" & Today().ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                              "WHERE " & vbCrLf & _
                              "Registro_Id = " & gridRegistro.Rows(i).Cells(1).Text()
                        End If
                    End If
                Else
                    If FinanceiroNovo Then
                        Dim TituloXHistorico As New Novo.TituloXHistorico()
                        TituloXHistorico.Acao = "BLOQUEAR"
                        TituloXHistorico.CodigoTitulo = gridRegistro.Rows(i).Cells(1).Text
                        TituloXHistorico.Data = DateTime.Now
                        TituloXHistorico.Historico = "BLOQUEIO DE TíTULO "
                        TituloXHistorico.Usuario = Session("ssNomeUsuario")
                        TituloXHistorico.Salvar()
                    Else
                        Sql = "UPDATE ContasAPagar " & vbCrLf & _
                         "   SET UsuarioLiberacao = ''" & vbCrLf & _
                         "WHERE " & vbCrLf & _
                         "Registro_Id = " & gridRegistro.Rows(i).Cells(1).Text()
                    End If

                End If

                If Sql.Length > 0 Then SqlArray.Add(Sql)

                i += 1
            End While

            If SqlArray.Count > 0 Then
                i = 0
                If Banco.GravaBanco(SqlArray) Then
                    LinkLiberar.Parent.Visible = False

                    MsgBox(Me.Page, "Registro(s) Atualizado(s) com Sucesso.", eTitulo.Sucess)

                    LimparCampos("")
                    Consultar_Registros()
                    TabContainer1.ActiveTabIndex = 0
                Else
                    While i < gridRegistro.Rows.Count
                        CType(gridRegistro.Rows(i).FindControl("imgSituacao"), ImageButton).ImageUrl = "~/images/erro.jpg"
                        i += 1
                    End While
                    MsgBox(Me.Page, "Erro ao Atualizar Registro(s): " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI.")
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro(s).")
        End If
    End Sub

    Protected Sub gridProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtProdutoXRegistros As New DataTable()
        dtProdutoXRegistros.Columns.Add("Registro", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
        dtProdutoXRegistros.Columns.Add("Empresa", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Cliente", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Pedido", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Historico", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Dolar", Type.GetType("System.Decimal"))
        dtProdutoXRegistros.Columns.Add("Valor", Type.GetType("System.Decimal"))

        txtValor.Text = gridProduto.SelectedRow.Cells(5).Text()

        Dim i As Integer = 0
        Dim j As Integer = 0

        While j < gridRegistro.Rows.Count
            If gridRegistro.Rows(j).Cells(6).Text() = gridProduto.SelectedRow.Cells(1).Text() Then
                Dim drProdutoXRegistros As DataRow = dtProdutoXRegistros.NewRow()
                drProdutoXRegistros("Registro") = gridRegistro.Rows(j).Cells(1).Text()
                drProdutoXRegistros("Vencimento") = gridRegistro.Rows(j).Cells(2).Text()
                drProdutoXRegistros("Empresa") = gridRegistro.Rows(j).Cells(3).Text()
                drProdutoXRegistros("Cliente") = gridRegistro.Rows(j).Cells(4).Text()
                drProdutoXRegistros("Pedido") = gridRegistro.Rows(j).Cells(5).Text()
                drProdutoXRegistros("NomeProduto") = gridRegistro.Rows(j).Cells(6).Text()
                drProdutoXRegistros("Historico") = gridRegistro.Rows(j).Cells(7).Text()
                drProdutoXRegistros("Dolar") = gridRegistro.Rows(j).Cells(8).Text()
                drProdutoXRegistros("Valor") = gridRegistro.Rows(j).Cells(9).Text()

                dtProdutoXRegistros.Rows.Add(drProdutoXRegistros)
            End If
            j += 1
        End While

        gridProdutoXRegistros.DataSource = dtProdutoXRegistros
        gridProdutoXRegistros.DataBind()

        j = 0
        Dim valorTotal As Decimal = 0
        While j < gridProdutoXRegistros.Rows.Count
            i = 0
            While i < gridRegistro.Rows.Count
                If gridRegistro.Rows(i).Cells(6).Text() = gridProduto.SelectedRow.Cells(1).Text() AndAlso gridRegistro.Rows(i).Cells(1).Text() = gridProdutoXRegistros.Rows(j).Cells(0).Text() Then
                    If CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = True Then
                        CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = True
                        valorTotal += CDec(gridRegistro.Rows(i).Cells(9).Text())
                    Else
                        CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = False
                    End If
                End If
                i += 1
            End While
            j += 1
        End While

        If CType(gridProduto.Rows(gridProduto.SelectedRow.RowIndex).FindControl("chkProduto"), CheckBox).Checked = True Then
            chkProdutoXRegistrosXTodos.Checked = True
        Else
            chkProdutoXRegistrosXTodos.Checked = False
        End If

        lblTotalProdutosXRegistros.Text = valorTotal.ToString("N2")
        chkProdutoXRegistrosXTodos.Enabled = True
        imgBuscarCliente.Enabled = True

        TabContainer1.ActiveTabIndex = 2
        '' CONFERIDO
    End Sub

    Protected Sub chkProduto_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim regProduto As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(regProduto.NamingContainer, GridViewRow)

        Dim j As Integer = 0
        While j < gridRegistro.Rows.Count
            If gridRegistro.Rows(j).Cells(6).Text() = gridProduto.Rows(row.RowIndex).Cells(1).Text() Then
                If regProduto.Checked Then
                    CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Checked = True
                Else
                    CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Checked = False
                End If
            End If
            CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Enabled = False
            j += 1
        End While
        '' CONFERIDO
    End Sub

    Protected Sub chkProdutoXRegistros_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim regProdutoXRegistro As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(regProdutoXRegistro.NamingContainer, GridViewRow)

        Dim j As Integer = 0
        While j < gridRegistro.Rows.Count
            If gridRegistro.Rows(j).Cells(1).Text() = gridProdutoXRegistros.Rows(row.RowIndex).Cells(0).Text() Then
                If regProdutoXRegistro.Checked Then
                    CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Checked = True
                Else
                    CType(gridProduto.Rows(gridProduto.SelectedRow.RowIndex).FindControl("chkProduto"), CheckBox).Checked = False
                    chkProdutoXRegistrosXTodos.Checked = False
                    CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Checked = False
                End If
            End If
            CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Enabled = False
            j += 1
        End While
    End Sub

    Protected Sub imgBuscarCliente_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim dtProdutoXRegistros As New DataTable()
        dtProdutoXRegistros.Columns.Add("Registro", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
        dtProdutoXRegistros.Columns.Add("Empresa", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Cliente", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Pedido", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Historico", Type.GetType("System.String"))
        dtProdutoXRegistros.Columns.Add("Dolar", Type.GetType("System.Decimal"))
        dtProdutoXRegistros.Columns.Add("Valor", Type.GetType("System.Decimal"))

        Dim j As Integer = 0
        j = 0
        While j < gridRegistro.Rows.Count
            If gridRegistro.Rows(j).Cells(6).Text() = gridProduto.SelectedRow.Cells(1).Text() Then
                If gridRegistro.Rows(j).Cells(4).Text().Contains(Trim(txtClienteXProduto.Text.ToUpper)) Then
                    Dim drProdutoXRegistros As DataRow = dtProdutoXRegistros.NewRow()
                    drProdutoXRegistros("Registro") = gridRegistro.Rows(j).Cells(1).Text()
                    drProdutoXRegistros("Vencimento") = gridRegistro.Rows(j).Cells(2).Text()
                    drProdutoXRegistros("Empresa") = gridRegistro.Rows(j).Cells(3).Text()
                    drProdutoXRegistros("Cliente") = gridRegistro.Rows(j).Cells(4).Text()
                    drProdutoXRegistros("Pedido") = gridRegistro.Rows(j).Cells(5).Text()
                    drProdutoXRegistros("NomeProduto") = gridRegistro.Rows(j).Cells(6).Text()
                    drProdutoXRegistros("Historico") = gridRegistro.Rows(j).Cells(7).Text()
                    drProdutoXRegistros("Dolar") = gridRegistro.Rows(j).Cells(8).Text()
                    drProdutoXRegistros("Valor") = gridRegistro.Rows(j).Cells(9).Text()

                    dtProdutoXRegistros.Rows.Add(drProdutoXRegistros)

                    CType(gridRegistro.Rows(j).FindControl("chkTitulos"), CheckBox).Checked = False
                End If
            End If
            j += 1
        End While

        CType(gridProduto.Rows(gridProduto.SelectedRow.RowIndex).FindControl("chkProduto"), CheckBox).Checked = False
        chkProdutoXRegistrosXTodos.Checked = False
        lblTotalProdutosXRegistros.Text = "0,00"

        gridProdutoXRegistros.DataSource = dtProdutoXRegistros
        gridProdutoXRegistros.DataBind()
    End Sub

    Protected Sub chkProdutoXRegistrosXTodos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim regTodos As CheckBox = CType(sender, CheckBox)

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim valorTotal As Decimal = 0

        While j < gridProdutoXRegistros.Rows.Count
            If regTodos.Checked Then
                CType(gridProduto.Rows(gridProduto.SelectedRow.RowIndex).FindControl("chkProduto"), CheckBox).Checked = True
                CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = True
                valorTotal += CDec(gridProdutoXRegistros.Rows(j).Cells(8).Text())

                i = 0
                While i < gridRegistro.Rows.Count
                    If gridRegistro.Rows(i).Cells(1).Text() = gridProdutoXRegistros.Rows(j).Cells(0).Text() Then
                        CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = True
                    End If
                    CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Enabled = False
                    i += 1
                End While
            Else
                CType(gridProduto.Rows(gridProduto.SelectedRow.RowIndex).FindControl("chkProduto"), CheckBox).Checked = False
                CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = False

                i = 0
                While i < gridRegistro.Rows.Count
                    If gridRegistro.Rows(i).Cells(1).Text() = gridProdutoXRegistros.Rows(j).Cells(0).Text() Then
                        CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = False
                    End If
                    CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Enabled = False
                    i += 1
                End While
            End If
            j += 1
        End While

        lblTotalProdutosXRegistros.Text = valorTotal.ToString("N2")
    End Sub

    Protected Sub txtValor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If txtValor.Text.Length = 0 OrElse txtValor.Text = "___.___.___.___,__" Then
            Dim i As Integer = 0
            Dim j As Integer = 0

            While j < gridProdutoXRegistros.Rows.Count
                CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = False

                i = 0
                While i < gridRegistro.Rows.Count
                    If gridRegistro.Rows(i).Cells(1).Text() = gridProdutoXRegistros.Rows(j).Cells(0).Text() Then
                        CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = False
                    End If
                    CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Enabled = False
                    i += 1
                End While
                j += 1
            End While
        Else
            If CDec(txtValor.Text) > CDec(gridProduto.SelectedRow.Cells(5).Text()) Then
                MsgBox(Me.Page, "Valor não pode ser maior do que falta a ser Liberado.")
            Else
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim valorTotal As Decimal = 0

                While j < gridProdutoXRegistros.Rows.Count
                    If (valorTotal + CDec(gridProdutoXRegistros.Rows(j).Cells(8).Text())) > CDec(txtValor.Text) Then
                        CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = False

                        i = 0
                        While i < gridRegistro.Rows.Count
                            If gridRegistro.Rows(i).Cells(1).Text() = gridProdutoXRegistros.Rows(j).Cells(0).Text() Then
                                CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = False
                            End If
                            CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Enabled = False
                            i += 1
                        End While
                    Else
                        valorTotal += gridProdutoXRegistros.Rows(j).Cells(8).Text()

                        CType(gridProdutoXRegistros.Rows(j).FindControl("chkProdutoXRegistros"), CheckBox).Checked = True

                        i = 0
                        While i < gridRegistro.Rows.Count
                            If gridRegistro.Rows(i).Cells(1).Text() = gridProdutoXRegistros.Rows(j).Cells(0).Text() Then
                                CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = True
                            End If
                            CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Enabled = False
                            i += 1
                        End While
                    End If
                    j += 1
                End While

                lblTotalProdutosXRegistros.Text = valorTotal.ToString("N2")
            End If
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AUTORIZACAODERETIRADA")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub LinkLiberar_Click1(sender As Object, e As EventArgs) Handles LinkLiberar.Click
        Try
            If Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then

                SqlArray.Clear()
                Dim i As Integer = 0
                While i < gridRegistro.Rows.Count
                    Sql = ""
                    If CType(gridRegistro.Rows(i).FindControl("chkTitulos"), CheckBox).Checked = True Then
                        If String.IsNullOrWhiteSpace(Server.HtmlDecode(gridRegistro.Rows(i).Cells(12).Text)) Then
                            If FinanceiroNovo Then
                                Dim TituloXHistorico As New Novo.TituloXHistorico()
                                TituloXHistorico.Acao = "LIBERAR"
                                TituloXHistorico.CodigoTitulo = gridRegistro.Rows(i).Cells(1).Text
                                TituloXHistorico.Data = DateTime.Now
                                TituloXHistorico.Historico = "LIBERAÇÃO DE TíTULO "
                                TituloXHistorico.Usuario = Session("ssNomeUsuario")
                                TituloXHistorico.Salvar()
                            Else
                                Sql = "UPDATE ContasAPagar " & vbCrLf & _
                                  "   SET UsuarioLiberacao = '" & Session("ssNomeUsuario") & "', " & vbCrLf & _
                                  "       UsuarioLiberacaoData = '" & Today().ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                                  "WHERE " & vbCrLf & _
                                  "Registro_Id = " & gridRegistro.Rows(i).Cells(1).Text()
                            End If
                        End If
                    Else
                        If FinanceiroNovo Then
                            Dim TituloXHistorico As New Novo.TituloXHistorico()
                            TituloXHistorico.Acao = "BLOQUEAR"
                            TituloXHistorico.CodigoTitulo = gridRegistro.Rows(i).Cells(1).Text
                            TituloXHistorico.Data = DateTime.Now
                            TituloXHistorico.Historico = "BLOQUEIO DE TíTULO "
                            TituloXHistorico.Usuario = Session("ssNomeUsuario")
                            TituloXHistorico.Salvar()
                        Else
                            Sql = "UPDATE ContasAPagar " & vbCrLf & _
                             "   SET UsuarioLiberacao = ''" & vbCrLf & _
                             "WHERE " & vbCrLf & _
                             "Registro_Id = " & gridRegistro.Rows(i).Cells(1).Text()
                        End If

                    End If

                    If Sql.Length > 0 Then SqlArray.Add(Sql)

                    i += 1
                End While

                If SqlArray.Count > 0 Then
                    i = 0
                    If Banco.GravaBanco(SqlArray) Then
                        LinkLiberar.Parent.Visible = False

                        MsgBox(Me.Page, "Registro(s) Atualizado(s) com Sucesso.", eTitulo.Sucess)

                        LimparCampos("")
                        Consultar_Registros()
                        TabContainer1.ActiveTabIndex = 0
                    Else
                        While i < gridRegistro.Rows.Count
                            CType(gridRegistro.Rows(i).FindControl("imgSituacao"), ImageButton).ImageUrl = "~/images/erro.jpg"
                            i += 1
                        End While
                        MsgBox(Me.Page, "Erro ao Atualizar Registro(s): " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro(s).")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class