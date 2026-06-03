Imports NGS.Lib.Negocio

Public Class RelatorioDeEstoqueDeTerceiros
    Inherits BasePage

#Region "Properties"

    Private Property CodigoCliente() As String
        Get
            Return ViewState("CodigoCliente")
        End Get
        Set(ByVal value As String)
            ViewState("CodigoCliente") = value
        End Set
    End Property

    Private Property HID() As String
        Get
            Return ViewState("HID")
        End Get
        Set(ByVal value As String)
            ViewState("HID") = value
        End Set
    End Property

#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Custos)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeEstoqueDeTerceiros", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
                    txtDataAte.Text = DateTime.Now.ToString("dd/MM/yyyy")

                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/ApuracaoDeCustos.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscarClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscarClientes.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objCliente" & HID, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidarCampos() Then
                Dim ds As DataSet = getDataSet()

                If ds.Tables(0).Rows.Count > 0 Then

                    Dim cli As String = String.Empty

                    For Each dr As DataRow In ds.Tables(0).Rows
                        cli = Funcoes.FormatarCpfCnpj(dr("Cliente_id"))
                        dr("Cliente_id") = cli
                    Next

                    Funcoes.BindReport(Me.Page, ds, "Cr_EstoqueDeTerceiros", eExportType.PDF, getParametros())
                Else
                    MsgBox(Me.Page, "Sem informações para essa seleção.", eTitulo.Info)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeEstoqueDeTerceiros")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Function getParametros() As Dictionary(Of String, Object)
        Dim param As New Dictionary(Of String, Object)
        param.Add("EmPoder", ddlEmPoder.SelectedValue)
        param.Add("MostrarPedido", Not chkMostrarPedidos.Checked)
        param.Add("paramConsulta", "")

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim obj As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param("paramConsulta") = Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & IIf(chkConsolidarEmpresa.Checked, " - Consolidado", "") & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            Dim obj As New Cliente(CodigoCliente.Split("-")(0), CodigoCliente.Split("-")(1))
            param("paramConsulta") &= Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & IIf(chkConsolidarCliente.Checked, " - Consolidado", "") & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim arr As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("", "Prd.Produto_id")
            param("paramConsulta") &= arr(1)
        End If

        If Not String.IsNullOrWhiteSpace(ddlEmPoder.SelectedValue) Then
            param("paramConsulta") &= ddlEmPoder.SelectedValue & " - " & IIf(chkMostrarPedidos.Checked, "Mostar Pedidos.", "") & " Data até: " & txtDataAte.Text & vbCrLf
        End If
        Return param
    End Function

    Private Sub LimparCampos()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        chkConsolidarEmpresa.Checked = False
        chkConsolidarCliente.Checked = False
        txtCliente.Text = String.Empty
        CodigoCliente = String.Empty
        HID = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID)
        ucSelecaoProduto.Limpar()
        ddlEmPoder.SelectedValue = String.Empty
        chkMostrarPedidos.Checked = False
        txtDataAte.Text = DateTime.Now.ToString("dd/MM/yyyy")

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a Empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataAte.Text) OrElse Not IsDate(txtDataAte.Text) Then
            MsgBox(Me.Page, "Informe a data.")
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objCliente" & HID) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(obj)
                txtCliente.Text = itemCliente.Text
                CodigoCliente = itemCliente.Value
                Session.Remove("objCliente" & HID)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String
        sql = "Declare @movimento nvarchar(10);" & vbCrLf & _
              "Set @movimento = '" & CDate(txtDataAte.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "Select sb.Empresa_id," & vbCrLf & _
              "       Sb.EndEmpresa_id," & vbCrLf & _
              "       Emp.Nome as NomeEmpresa," & vbCrLf & _
              "       Emp.Complemento," & vbCrLf & _
              "       Emp.Cidade," & vbCrLf & _
              "       Emp.Estado," & vbCrLf & _
              "       sb.Classificacao," & vbCrLf & _
              "       sb.Cliente_id," & vbCrLf & _
              "       sb.EndCliente_Id," & vbCrLf & _
              "       cli.Nome as NomeCliente," & vbCrLf

        'If ChkSomenteValor.Checked Then
        '    sql &= "          '0' as Produto_id," & vbCrLf & _
        '           "          'Sem Produto - Somente Valor' as NomeProduto," & vbCrLf
        'Else
        sql &= "          sb.Produto_id," & vbCrLf &
               "          Prd.Nome as NomeProduto," & vbCrLf
        'End If

        If chkMostrarPedidos.Checked Then
            sql &= "          sb.Pedido," & vbCrLf
        End If

        sql &= "          cast(sum(sb.qtdeRazao) as decimal(18,4))  as QtdeRazao," & vbCrLf &
               "          cast(Sum(sb.Qtdefiscal) as decimal(18,4)) as QtdeFiscal," & vbCrLf &
               "          cast(sum(sb.qtdeRazao + sb.Qtdefiscal) as decimal(18,4)) as Qtde," & vbCrLf &
               "          cast(sum(sb.ValorRazao) as decimal(18,2)) as ValorRazao," & vbCrLf &
               "          cast(sum(sb.ValorFiscal) as decimal(18,2)) as ValorFiscal," & vbCrLf &
               "          cast(Sum(sb.ValorRazao + sb.ValorFiscal) as decimal(18,2)) as Valor" & vbCrLf &
               "     from (" & vbCrLf &
               "   		select nfi.Empresa_Id," & vbCrLf &
               "   		       nfi.EndEmpresa_id," & vbCrLf &
               "   			   case" & vbCrLf &
               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf &
               "   				  when SO.ProdutoDeTerceiro =  1                                                                    then 'TERCEIROS EM NOSSO PODER'" & vbCrLf &
               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf &
               "   			   end as Classificacao," & vbCrLf &
               "   			   case" & vbCrLf &
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf &
               "   				   then nf.destino" & vbCrLf &
               "   				   else nfi.cliente_id" & vbCrLf &
               "   			   end Cliente_id," & vbCrLf &
               "   			   case" & vbCrLf &
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf &
               "   				   then nf.Enddestino" & vbCrLf &
               "   				   else nfi.Endcliente_id" & vbCrLf &
               "   			   end EndCliente_id," & vbCrLf &
               "   			   nfi.Produto_Id," & vbCrLf &
               "   			   isnull(nf.pedido,0) as Pedido," & vbCrLf &
               "                  convert(numeric(18,4),0) as QtdeRazao," & vbCrLf &
               "   			   SUM(Case" & vbCrLf &
               "   					 When So.Devolucao = 'S'" & vbCrLf &
               "   					   then nfi.QuantidadeFiscal * -1" & vbCrLf &
               "   					   else nfi.QuantidadeFiscal" & vbCrLf &
               "   				   end) QtdeFiscal," & vbCrLf &
               "               convert(numeric(18,4),0) as ValorRazao," & vbCrLf &
               "   			   SUM(Case" & vbCrLf &
               "   					 When So.Devolucao = 'S'" & vbCrLf &
               "   					   then nfe.valor * -1" & vbCrLf &
               "   					   else nfe.valor" & vbCrLf &
               "   				   end) as valorFiscal" & vbCrLf &
               "   		  from notasfiscais NF" & vbCrLf &
               "   		 inner join Notasfiscaisxitens NFI" & vbCrLf &
               "   			on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf &
               "   		   and NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
               "   		   and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
               "   		   and NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
               "   		   and NF.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf &
               "   		   and NF.Nota_id         = NFI.Nota_id" & vbCrLf &
               "   		   and NF.Serie_Id        = NFI.Serie_Id" & vbCrLf &
               "   		 Inner Join NotasFiscaisxEncargos Nfe" & vbCrLf &
               "   			on Nfe.Empresa_id      = NFI.Empresa_Id" & vbCrLf &
               "   		   and Nfe.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
               "   		   and Nfe.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
               "   		   and Nfe.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
               "   		   and Nfe.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf &
               "   		   and Nfe.Nota_id         = NFI.Nota_id" & vbCrLf &
               "   		   and Nfe.Serie_Id        = NFI.Serie_Id" & vbCrLf &
               "   		   and Nfe.Produto_id      = NFI.Produto_id" & vbCrLf &
               "   		   and Nfe.CFOP_Id         = NFI.CFOP_Id" & vbCrLf &
               "   		   and Nfe.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf &
               "   		   and Nfe.Encargo_Id      = 'LIQUIDO'" & vbCrLf &
               "   		 Inner join produtos p" & vbCrLf &
               "   			on NFI.Produto_id = P.Produto_id" & vbCrLf &
               "   		 Inner Join Suboperacoes SO" & vbCrLf &
               "   			on SO.Operacao_Id        = NFI.Operacao" & vbCrLf &
               "   		   and SO.Suboperacoes_Id    = NFI.Suboperacao" & vbCrLf &
               "   		 where (SO.produtodeterceiro = 1 or SO.deposito = 'S')" & vbCrLf &
               "   		   and (SO.EstoqueFisico = 'S' or SO.EstoqueFiscal = 'S' or SO.Devolucao = 'S' or (NFI.QuantidadeFiscal = 0 and NFI.Unitario = 0 and NFI.Valor > 0))" & vbCrLf &
               "   		   and SO.Operacao_id <> 80" & vbCrLf &
               "   		   and nf.movimento  <= @movimento" & vbCrLf
        '"   		   and year(nf.movimento) > 2010" & vbCrLf & _
        '"   		   and nf.situacao    = 1" & vbCrLf & _
        '"           and not (nf.empresa_id = '04854422000185' and nf.pedido in(1686,598,101000679,876,1746,4521,7293,595,7153,6425,1533))" & vbCrLf & _
        '"           and not (nf.empresa_id = '04854422000266' and nf.pedido in(102003060,102002278,102002277,102002071))" & vbCrLf & _
        '"           and not (nf.empresa_id = '04854422000428' and nf.pedido in(104000219,104000004,104000339))"

        If ddlEmPoder.SelectedValue = "Em Nosso Poder" Then
            sql &= " and SO.ProdutoDeTerceiro =  1" & vbCrLf
        ElseIf ddlEmPoder.SelectedValue = "Poder de Terceiros" Then
            sql &= " and So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S')" & vbCrLf
        End If

        sql &= "   		  group by nfi.Empresa_Id," & vbCrLf & _
               "   		           nfi.EndEmpresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   				  when SO.ProdutoDeTerceiro = 1                                                                     then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id, 8) = left(nfi.Cliente_id, 8)" & vbCrLf & _
               "   				   then nf.destino" & vbCrLf & _
               "   				   else nfi.cliente_id" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
               "   				   then nf.Enddestino" & vbCrLf & _
               "   				   else nfi.Endcliente_id" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "               nf.pedido," & vbCrLf & _
               "               nfi.Produto_Id" & vbCrLf & _
               "        union all" & vbCrLf & _
               "   		Select R.Empresa_Id," & vbCrLf & _
               "   			   R.Endempresa_id," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
               "   			   end Classificacao," & vbCrLf & _
               "   			   Case" & vbCrLf & _
               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
               "                   then R.Deposito" & vbCrLf & _
               "                   else R.Cliente_id" & vbCrLf & _
               "               end Cliente_id," & vbCrLf & _
               "   			   Case" & vbCrLf & _
               "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
               "                   then R.EndDeposito" & vbCrLf & _
               "                   else R.EndCliente_id" & vbCrLf & _
               "               end EndCliente_id," & vbCrLf & _
               "   			   R.Produto," & vbCrLf & _
               "   			   isnull(R.Pedido,0) as Pedido," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then Isnull(R.CreditoQuantidade, 0.0000) - Isnull(R.DebitoQuantidade, 0.0000)" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoQuantidade, 0.0000)  - Isnull(R.CreditoQuantidade, 0.0000)" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "   			   convert(numeric(18,4),0)," & vbCrLf & _
               "   			   case" & vbCrLf & _
               "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then isnull(R.CreditoOficial, 0.00) - isnull(R.DebitoOficial, 0.00)" & vbCrLf & _
               "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoOficial, 0.00)  - isnull(R.CreditoOficial, 0.00)" & vbCrLf & _
               "   			   end," & vbCrLf & _
               "               convert(numeric(18,4),0)" & vbCrLf & _
               "   		  from Razao R" & vbCrLf & _
               "   		 inner join ClientesxEmpresas CxE" & vbCrLf & _
               "   			on R.Empresa_id    = CxE.Empresa_Id" & vbCrLf & _
               "   		   and R.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
               "   		   and (left(R.Conta_Id,7) = CxE.ContaEstoqueEmNossoPoder or Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) = Left(cxe.ContaestoqueemPoderdeTerceiros,5))" & vbCrLf & _
               "   		 inner join Clientes C" & vbCrLf & _
               "   			on R.Cliente_id    = C.Cliente_id" & vbCrLf & _
               "   		   and R.EndCliente_id = C.Endereco_id" & vbCrLf & _
               "   		 where lote_id not in (9,10,11)" & vbCrLf & _
               "           and R.Movimento_Id <= @movimento" & vbCrLf & _
               "           and year(R.Movimento_Id) > 2010 " & vbCrLf & _
               "   		   and len(isnull(produto,'')) > 0" & vbCrLf

        If ddlEmPoder.SelectedValue = "Em Nosso Poder" Then
            sql &= " and R.Conta_id = CxE.ContaEstoqueEmNossoPoder" & vbCrLf
        ElseIf ddlEmPoder.SelectedValue = "Poder de Terceiros" Then
            sql &= " and Left(R.Conta_id,5) = left(CxE.ContaEstoqueEmPoderDeTerceiros,5)" & vbCrLf
        End If

        sql &= "   		) as Sb" & vbCrLf & _
               "     inner join Clientes Emp" & vbCrLf & _
               "        on Sb.Empresa_id    = Emp.Cliente_id" & vbCrLf & _
               "       and Sb.EndEmpresa_id = Emp.Endereco_id" & vbCrLf & _
               "     inner join Clientes Cli" & vbCrLf & _
               "        on Sb.cliente_id    = Cli.Cliente_id" & vbCrLf & _
               "       and Sb.EndCliente_id = Cli.Endereco_id" & vbCrLf & _
               "     inner Join Produtos Prd" & vbCrLf & _
               "        on Prd.Produto_id = sb.Produto_id" & vbCrLf & _
               "     Where Prd.TipoDoItem Not in(7,8,9,10,99)" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            sql &= " and " & IIf(chkConsolidarEmpresa.Checked, "left(sb.Empresa_id, 8) = '" & Left(ddlEmpresa.SelectedValue, 8) & "'", "sb.Empresa_id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'")
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            sql &= " and " & IIf(chkConsolidarCliente.Checked, "left(sb.Cliente_id, 8) = '" & Left(CodigoCliente, 8) & "'", "sb.Cliente_id = '" & CodigoCliente.Split("-")(0) & "'")
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("", "Prd.Produto_id")
            sql &= " AND " & RetornoProdutos(0)
        End If

        sql &= "    Group by sb.Empresa_id," & vbCrLf & _
               "          Sb.EndEmpresa_id," & vbCrLf & _
               "          Emp.Nome," & vbCrLf & _
               "          Emp.Complemento," & vbCrLf & _
               "          Emp.Cidade," & vbCrLf & _
               "          Emp.Estado," & vbCrLf & _
               "          sb.Classificacao," & vbCrLf & _
               "          sb.Cliente_id," & vbCrLf & _
               "          sb.EndCliente_Id," & vbCrLf & _
               "          cli.Nome" & vbCrLf

        'If Not ChkSomenteValor.Checked Then
        sql &= "         ,sb.Produto_id," & vbCrLf &
                "          Prd.Nome" & vbCrLf
        'End If

        If chkMostrarPedidos.Checked Then
            sql &= "     ,sb.Pedido" & vbCrLf
        End If

        sql &= "    having SUM(sb.QtdeRazao + sb.QtdeFiscal)   <> 0  Or Sum(sb.ValorRazao + sb.ValorFiscal) <> 0" & vbCrLf & _
               "    order by cli.Nome" & IIf(Not ChkSomenteValor.Checked, ", sb.Produto_id", "") & vbCrLf


        Return Banco.ConsultaDataSet(sql, "EstoqueDeTerceiros")
    End Function

#End Region

End Class