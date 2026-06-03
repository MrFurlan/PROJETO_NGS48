Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Public Class LiberaPedidoVirtual
    Inherits BasePage

    Private objPedido As [Lib].Negocio.Pedido
    Private sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("LiberaPedidoVirtual", "ACESSAR") AndAlso FinanceiroVirtual Then
                CargaUnidadeDeNegocio()
                ddl.Carregar(ddlCondPagPed, CarregarDDL.Tabela.CondicaoDePagamento)

                Limpar(True)

                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
        If (Session("RowIndex" & HID.Value) IsNot Nothing) Then
            gridConsulta.SelectedIndex = Session("RowIndex")
            Dim imgPedido As ImageButton = CType(gridConsulta.Rows(CInt(Session("RowIndex"))).FindControl("imgSelecionar"), ImageButton)
            imgPedido.Focus()
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objClienteLibPed" & HID.Value.ToString) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteLibPed" & HID.Value), [Lib].Negocio.Cliente))
                txtCodigoCliente.Value = itemCliente.Value
                txtCliente.Text = itemCliente.Text
                Session.Remove("objClienteLibPed" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar(ByVal geral As Boolean)
        lnkUmNovo.Parent.Visible = False
        lnkVirtualizar.Parent.Visible = False
        lnkVincular.Parent.Visible = False

        If geral Then
            txtCodigoCliente.Value = String.Empty
            txtCliente.Text = String.Empty
            cmdArquivoDeSaida.Visible = False
        End If

        txtCodigoPedido.text = String.Empty

        txtCnpjDaEmpresa.Text = String.Empty
        txtNomeDaEmpresa.Text = String.Empty
        txtCidadeDaEmpresa.Text = String.Empty
        txtCnpjCliente.Text = String.Empty
        txtNomeDoCliente.Text = String.Empty
        txtCidadeCliente.Text = String.Empty
        txtDeposito.Text = String.Empty
        txtNomeDoDeposito.Text = String.Empty
        txtCidadeDeposito.Text = String.Empty
        txtPedido.Text = String.Empty
        txtDataEntrega.Text = String.Empty
        txtOperacao.Text = String.Empty
        txtMoeda.Text = String.Empty
        txtNaturezaDaOperacao.Text = String.Empty
        txtTotalNotas.Text = String.Empty
        txtTotalTitulos.Text = String.Empty

        gridFinanceiro.DataSource = Nothing
        gridFinanceiro.DataBind()

        GridNotas.DataSource = Nothing
        GridNotas.DataBind()

        Session.Remove("RowIndex" & HID.Value)
        Session.Remove("objClienteLibPed" & HID.Value)
        Session.Remove("objPedido" & HID.Value)
        Session.Remove("objPedVencimentos" & HID.Value)

        gridConsulta.DataSource = Nothing
        gridConsulta.DataBind()

        ddlCondPagPed.SelectedIndex = 0
        txtDataCondPagParcela.Text = String.Empty
        txtPedidoTotal.Text = String.Empty
        txtPedidoTotalPago.Text = String.Empty
        txtPedidoSaldo.Text = String.Empty

        gridParcelas.DataSource = Nothing
        gridParcelas.DataBind()

        GridNotasXTitulos.DataSource = Nothing
        GridNotasXTitulos.DataBind()

        txtCodigoParcela.Text = String.Empty
        txtDataVencParcela.Text = String.Empty
        txtValorParcela.Text = String.Empty

        TabParcelamento.Visible = False

        HID.Value = (New Random).Next
        TabContainer1.ActiveTabIndex = 0
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub SalvarPedido()
        Session("objPedido" & HID.Value) = objPedido
    End Sub

    Private Sub RecuperarPedido()
        objPedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
    End Sub

    Private Sub ConsultaRegistros()
        If Funcoes.VerificaPermissao("LiberaPedidoVirtual", "LEITURA") Then
            sql = "select p.Empresa_id, p.EndEmpresa_Id, cliEmp.Nome, cliEmp.Cidade, p.Pedido_Id, P.PedidoEfetivo, p.Cliente, p.EndCliente, cli.Nome as NomeCliente,  " & vbCrLf & _
                "		p.DataPedido, p.Moeda, " & vbCrLf & _
                "Sum(pXi.Quantidade) AS Quantidade, " & vbCrLf & _
                "sum(case  " & vbCrLf & _
                "	when p.Moeda = 1 " & vbCrLf & _
                "		then pXi.TotalOficial " & vbCrLf & _
                "		else pXi.TotalMoeda " & vbCrLf & _
                "	end) ValorPedido, " & vbCrLf & _
                "sum(isnull(fin.finanAberto,0)) AS valoraberto," & vbCrLf & _
                "sum(isnull(fin.finanBaixado,0)) AS valorbaixado" & vbCrLf & _
                " from Pedidos p  " & vbCrLf & _
                "     inner join (Select pedI.Empresa_id, pedI.EndEmpresa_Id, pedI.Pedido_Id," & vbCrLf & _
                "                     sum(case    " & vbCrLf & _
                " 						  when pedI.TipoDeLancamento = 'E'   " & vbCrLf & _
                " 							then pedI.Quantidade * -1   " & vbCrLf & _
                " 							else pedI.Quantidade   " & vbCrLf & _
                " 						end) as Quantidade,   " & vbCrLf & _
                " 				    sum(case    " & vbCrLf & _
                " 						  when pedI.TipoDeLancamento = 'E'   " & vbCrLf & _
                " 							then pedI.TotalOficial * -1    " & vbCrLf & _
                " 							else pedI.TotalOficial   " & vbCrLf & _
                " 						end) as TotalOficial,   " & vbCrLf & _
                " 				    sum(case    " & vbCrLf & _
                " 						  when pedI.TipoDeLancamento = 'E'   " & vbCrLf & _
                " 							then pedI.TotalMoeda * -1    " & vbCrLf & _
                " 							else pedI.TotalMoeda   " & vbCrLf & _
                " 						end) as TotalMoeda   " & vbCrLf & _
                "				from PedidoXItemxLancamento pedI  " & vbCrLf & _
                "					inner join Pedidos ped  " & vbCrLf & _
                "							on ped.Empresa_Id     = pedI.Empresa_Id  " & vbCrLf & _
                "							and ped.EndEmpresa_Id = pedI.EndEmpresa_Id  " & vbCrLf & _
                "							and ped.Pedido_Id     = pedI.Pedido_Id  " & vbCrLf & _
                "				where ped.Situacao = 1  " & vbCrLf
            If ddlEmpresa.SelectedIndex > 0 Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                sql &= "				  and ped.Empresa_id = '" & Empresa(0) & "' and ped.EndEmpresa_id = " & Empresa(1) & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoPedido.Text) Then
                sql &= "				  and ped.Pedido_Id = " & txtCodigoPedido.Text & vbCrLf
            End If

            sql &= "				group by pedI.Empresa_Id, pedI.EndEmpresa_Id, pedI.Pedido_Id) pXi  " & vbCrLf & _
                "			on pXi.Empresa_Id = p.Empresa_Id  " & vbCrLf & _
                "			and pXi.EndEmpresa_Id = p.EndEmpresa_Id  " & vbCrLf & _
                "			and pXi.Pedido_Id = p.Pedido_Id  " & vbCrLf & _
                "	 left join PedidoXParcela pXp  " & vbCrLf & _
                "			on pXp.Empresa_Id     = p.Empresa_Id  " & vbCrLf & _
                "			and pXp.EndEmpresa_Id = p.EndEmpresa_Id  " & vbCrLf & _
                "			and pXp.Pedido_Id     = p.Pedido_Id  " & vbCrLf & _
                "    inner join (Select n.Empresa_id, n.EndEmpresa_Id, n.Pedido, isnull(n.NFG,0) as NFG" & vbCrLf & _
                "               from NotasFiscais n" & vbCrLf & _
                "               where n.Situacao = 1" & vbCrLf
            If ddlEmpresa.SelectedIndex > 0 Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                sql &= "				  and n.Empresa_id = '" & Empresa(0) & "' and n.EndEmpresa_id = " & Empresa(1) & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoPedido.Text) Then
                sql &= "				  and n.Pedido = " & txtCodigoPedido.Text & vbCrLf
            End If

            sql &= "               group by n.Empresa_Id, n.EndEmpresa_Id, n.Pedido, isnull(n.NFG,0)) NF " & vbCrLf & _
                "           on NF.Empresa_Id     = p.Empresa_Id" & vbCrLf & _
                "           and NF.EndEmpresa_Id = p.EndEmpresa_Id" & vbCrLf & _
                "           and NF.Pedido        = p.Pedido_Id " & vbCrLf & _
                "	left join (Select cp.EmpresaPedido, cp.EndEmpresaPedido, cp.Pedido, isnull(cXp.BaixaAdiantamento,0) AS BaixaAdiantamento," & vbCrLf & _
                "				sum(case" & vbCrLf & _
                "					when not cp.provisao = 1" & vbCrLf & _
                "						then cp.valordodocumento" & vbCrLf & _
                "						else 0" & vbCrLf & _
                "					end) as finanAberto," & vbCrLf & _
                "				sum(case" & vbCrLf & _
                "					when cp.provisao = 1" & vbCrLf & _
                "						then cp.valordodocumento" & vbCrLf & _
                "						else 0" & vbCrLf & _
                "					end) as finanBaixado" & vbCrLf & _
                "				from " & IIf(rdEntrada.Checked, "ContasAPagar", "ContasAReceber") & " cp" & vbCrLf & _
                "					inner join ComprasXProdutos cXp" & vbCrLf & _
                "							on cXp.Produto_Id = cp.Carteira" & vbCrLf & _
                "               where cp.Situacao = 1 and isnull(cp.PedidoFixacao,0) = 0" & vbCrLf
            If ddlEmpresa.SelectedIndex > 0 Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                sql &= "				  and cp.EmpresaPedido = '" & Empresa(0) & "' and cp.EndEmpresaPedido = " & Empresa(1) & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoPedido.text) Then
                sql &= "				  and cp.Pedido = " & txtCodigoPedido.Text & vbCrLf
            End If

            sql &= "				group by cp.EmpresaPedido, cp.EndEmpresaPedido, cp.Pedido, isnull(cXp.BaixaAdiantamento,0)) fin" & vbCrLf & _
                "			on fin.EmpresaPedido = p.Empresa_Id" & vbCrLf & _
                "			and fin.EndEmpresaPedido = p.EndEmpresa_Id" & vbCrLf & _
                "			and fin.Pedido           = p.Pedido_Id" & vbCrLf & _
                "	 inner join SubOperacoes so  " & vbCrLf & _
                "	    	 on so.Operacao_id     = p.Operacao  " & vbCrLf & _
                "		    and so.SubOperacoes_id = p.Suboperacao  " & vbCrLf & _
                "	 inner join Clientes cliEmp  " & vbCrLf & _
                "	    	 on cliEmp.Cliente_id  = p.Empresa_Id  " & vbCrLf & _
                "	    	and cliEmp.Endereco_id = p.EndEmpresa_Id  " & vbCrLf & _
                "	inner join Clientes cli  " & vbCrLf & _
                "	    	 on cli.Cliente_id  = p.Cliente  " & vbCrLf & _
                "	    	and cli.Endereco_id = p.EndCliente  " & vbCrLf & _
                " where p.Situacao = 1  " & vbCrLf & _
                "  and so.Financeiro = 'S' " & vbCrLf & _
                "  and isnull(pXp.Pedido_Id,0) = 0 " & vbCrLf & _
                "  and isnull(NF.NFG,0) = 0 " & vbCrLf

            If rdReal.Checked Then
                sql &= "  and p.Moeda = 1" & vbCrLf
            Else
                sql &= "  and p.Moeda = 3" & vbCrLf
            End If

            If rdEntrada.Checked Then
                sql &= "  and so.EntradaSaida = 'E'" & vbCrLf
            Else
                sql &= "  and so.EntradaSaida = 'S'" & vbCrLf
            End If

            If ddlEmpresa.SelectedIndex > 0 Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                sql &= "  and p.Empresa_id = '" & Empresa(0) & "' " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoPedido.Text) Then
                sql &= "  and p.Pedido_id = " & txtCodigoPedido.Text & vbCrLf
            End If

            If txtCliente.Text.Length > 0 Then
                Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                sql &= "  and p.Cliente    = '" & Cliente(0) & "'" & vbCrLf & _
                        "  and p.EndCliente = " & Cliente(1) & vbCrLf
            End If

            If String.IsNullOrWhiteSpace(txtCodigoPedido.Text) Then
                sql &= " and p.Datapedido between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
            End If

            ' apenas abertos
            If rdAberto.Checked Then sql &= "  and isnull(fin.finanBaixado,0) = 0 and isnull(fin.finanAberto,0) > 0 " & vbCrLf

            sql &= "group by p.Empresa_id, p.EndEmpresa_Id, cliEmp.Nome, cliEmp.Cidade, p.Pedido_Id, P.PedidoEfetivo, p.Cliente, p.EndCliente, cli.Nome," & vbCrLf & _
                   "p.DataPedido, p.Moeda" & vbCrLf
            sql &= "order by p.Datapedido " & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "ConsultaPedidos")

            gridConsulta.DataSource = Nothing
            gridConsulta.DataBind()

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum Registro Encontrado")
                Exit Sub
            End If

            Dim dtPedidos As New DataTable("listPedidos")
            dtPedidos.Columns.Add("Empresa", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Cliente", Type.GetType("System.String"))
            dtPedidos.Columns.Add("NomeCliente", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Pedido", Type.GetType("System.String"))
            dtPedidos.Columns.Add("PedidoEfetivo", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Data", Type.GetType("System.DateTime"))
            dtPedidos.Columns.Add("Moeda", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("ValorPedido", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("valoraberto", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("valorbaixado", Type.GetType("System.Decimal"))

            For Each dr As DataRow In ds.Tables(0).Rows
                Dim drPedido As DataRow = dtPedidos.NewRow()
                drPedido("Empresa") = dr("Empresa_Id") & "-" & Left(dr("Nome"), 20) & "/" & dr("Cidade")
                drPedido("Cliente") = dr("Cliente") & "-" & dr("EndCliente")
                drPedido("NomeCliente") = dr("NomeCliente")
                drPedido("Pedido") = dr("Pedido_Id")
                drPedido("PedidoEfetivo") = dr("PedidoEfetivo")
                drPedido("Data") = dr("DataPedido")
                If dr("Moeda") = 1 Then
                    drPedido("Moeda") = "R$"
                Else
                    drPedido("Moeda") = "U$"
                End If
                drPedido("Quantidade") = dr("Quantidade")
                drPedido("ValorPedido") = dr("ValorPedido")
                drPedido("valoraberto") = dr("valoraberto")
                drPedido("valorbaixado") = dr("valorbaixado")
                dtPedidos.Rows.Add(drPedido)
            Next

            gridConsulta.DataSource = dtPedidos
            gridConsulta.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidadeDeNegocio.SelectedIndex > 0 Then
                CargaEmpresa()
            Else
                ddlEmpresa.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClienteLibPed" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedido.Click
        Try
            If txtPedido.Text.Length = 0 Then
                MsgBox(Me.Page, "Consulte o Registro para visualização do Extrato")
            ElseIf txtCnpjDaEmpresa.Text.Length = 0 Then
                MsgBox(Me.Page, "Empresa do Registro não encontrada")
            ElseIf txtCnpjCliente.Text.Length = 0 Then
                MsgBox(Me.Page, "Cliente do Registro não encontrado")
            Else
                RecuperarPedido()
                Extrato.Emitir(Me.Page, False, objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, "T", objPedido.Codigo)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkSelecionar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnkPedido As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnkPedido.NamingContainer, GridViewRow)
            Session("RowIndex" & HID.Value) = row.RowIndex
            gridConsulta.SelectedIndex = row.RowIndex

            Dim strPedido() As String = gridConsulta.Rows(row.RowIndex).Cells(1).Text.Split("-")

            objPedido = New [Lib].Negocio.Pedido(strPedido(0), 0, gridConsulta.Rows(row.RowIndex).Cells(4).Text)

            If objPedido.TemFixacoes Then
                MsgBox(Me.Page, "Pedido com Fixação não pode ser Virtualizado.")
                Exit Sub
            End If

            If objPedido.Vencimentos.Count > 0 AndAlso objPedido.Vencimentos.Where(Function(s) s.RegistroMestre > 0).Count > 0 Then
                MsgBox(Me.Page, "Pedido com Título Agrupado ainda não está liberado.")
                Exit Sub
            End If

            Dim vDeducoesP As Decimal = 0
            Dim vDescontosP As Decimal = 0
            Dim vAcrescimosP As Decimal = 0
            Dim vJurosP As Decimal = 0

            Dim vDeducoesR As Decimal = 0
            Dim vDescontosR As Decimal = 0
            Dim vAcrescimosR As Decimal = 0
            Dim vJurosR As Decimal = 0

            Dim vDeducoesMoeda As Decimal = 0
            Dim vDescontosMoeda As Decimal = 0
            Dim vAcrescimosMoeda As Decimal = 0
            Dim vJurosMoeda As Decimal = 0

            Dim totalDeTitulosP As Integer = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "P").Count
            Dim totalDeTitulosR As Integer = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "R").Count

            If objPedido.Vencimentos.Count > 0 Then
                vDeducoesP = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "P").Sum(Function(s) s.DeducoesOficial)
                vDescontosP = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "P").Sum(Function(s) s.DescontosOficial)
                vAcrescimosP = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "P").Sum(Function(s) s.AcrescimosOficial)
                vJurosP = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "P").Sum(Function(s) s.JurosOficial)

                vDeducoesR = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "R").Sum(Function(s) s.DeducoesOficial)
                vDescontosR = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "R").Sum(Function(s) s.DescontosOficial)
                vAcrescimosR = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "R").Sum(Function(s) s.AcrescimosOficial)
                vJurosR = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa And s.TipoFinanceiro = "R").Sum(Function(s) s.JurosOficial)

                vDeducoesMoeda = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa).Sum(Function(s) s.DeducoesMoeda)
                vDescontosMoeda = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa).Sum(Function(s) s.DescontosMoeda)
                vAcrescimosMoeda = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa).Sum(Function(s) s.AcrescimosMoeda)
                vJurosMoeda = objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa).Sum(Function(s) s.JurosMoeda)
            End If

            Dim dtTitulos As New DataTable("listTitulos")
            dtTitulos.Columns.Add("Provisao", Type.GetType("System.String"))
            dtTitulos.Columns.Add("TipoFinanceiro", Type.GetType("System.String"))
            dtTitulos.Columns.Add("Codigo", Type.GetType("System.String"))
            dtTitulos.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
            dtTitulos.Columns.Add("DataBaixa", Type.GetType("System.DateTime"))
            dtTitulos.Columns.Add("Moeda", Type.GetType("System.String"))
            dtTitulos.Columns.Add("DocumentoOficial", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("DeducaoOficial", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("DescontoOficial", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("AcrescimoOficial", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("JurosOficial", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("LiquidoOficial", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("DocumentoMoeda", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("DeducaoMoeda", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("DescontoMoeda", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("AcrescimoMoeda", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("JurosMoeda", Type.GetType("System.Decimal"))
            dtTitulos.Columns.Add("LiquidoMoeda", Type.GetType("System.Decimal"))

            For Each parcela In objPedido.Vencimentos.ToList.OrderBy(Function(s) s.Provisao)
                Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(parcela.Carteira)
                If Not objCarteira.isAdiantamento Then
                    Dim drTitulo As DataRow = dtTitulos.NewRow()
                    drTitulo("Provisao") = parcela.Provisao.ToInt32
                    drTitulo("TipoFinanceiro") = parcela.TipoFinanceiro
                    drTitulo("Codigo") = parcela.Codigo
                    drTitulo("Vencimento") = parcela.DataProrrogacao
                    drTitulo("DataBaixa") = parcela.DataBaixa

                    If parcela.CodigoMoeda = 1 Then
                        drTitulo("Moeda") = "R$"
                    Else
                        drTitulo("Moeda") = "U$"
                    End If

                    drTitulo("DocumentoOficial") = parcela.ValorDocumentoOficial
                    drTitulo("DeducaoOficial") = parcela.DeducoesOficial
                    drTitulo("DescontoOficial") = parcela.DescontosOficial
                    drTitulo("AcrescimoOficial") = parcela.AcrescimosOficial
                    drTitulo("JurosOficial") = parcela.JurosOficial
                    drTitulo("LiquidoOficial") = parcela.ValorLiquidoOficial
                    drTitulo("DocumentoMoeda") = parcela.ValorDocumentoMoeda
                    drTitulo("DeducaoMoeda") = parcela.DeducoesMoeda
                    drTitulo("DescontoMoeda") = parcela.DescontosMoeda
                    drTitulo("AcrescimoMoeda") = parcela.AcrescimosMoeda
                    drTitulo("JurosMoeda") = parcela.JurosMoeda
                    drTitulo("LiquidoMoeda") = parcela.ValorLiquidoMoeda

                    dtTitulos.Rows.Add(drTitulo)
                End If
            Next

            gridFinanceiro.DataSource = dtTitulos
            gridFinanceiro.DataBind()

            If objPedido.Vencimentos.Count > 0 Then
                Dim vlrtitulos As Decimal = dtTitulos.Compute("SUM(DocumentoOficial)", "")
                txtTotalTitulos.Text = vlrtitulos.ToString("N2")
            End If

            sql = "SELECT DISTINCT NF.Movimento, " & vbCrLf & _
                    "            NF.Empresa_id, " & vbCrLf & _
                    "            NF.EndEmpresa_id, " & vbCrLf & _
                    "            NF.Cliente_Id AS Cliente, " & vbCrLf & _
                    "            NF.EndCliente_Id AS EndCliente, " & vbCrLf & _
                    "            C.Nome as NomeCliente, " & vbCrLf & _
                    "            NF.EntradaSaida_Id AS ES, " & vbCrLf & _
                    "            NF.Serie_Id AS Serie, " & vbCrLf & _
                    "            NF.Nota_Id AS Nota, " & vbCrLf & _
                    "            NF.Operacao, " & vbCrLf & _
                    "            NF.SubOperacao, " & vbCrLf & _
                    "            SUM(NE.Valor) AS Valor " & vbCrLf & _
                    "FROM NotasFiscaisXEncargos NE " & vbCrLf & _
                    "	INNER JOIN NotasFiscais NF " & vbCrLf & _
                    "			ON NF.Empresa_Id       = NE.Empresa_Id " & vbCrLf & _
                    "			AND NF.EndEmpresa_Id   = NE.EndEmpresa_Id " & vbCrLf & _
                    "			AND NF.Cliente_Id      = NE.Cliente_Id " & vbCrLf & _
                    "			AND NF.EndCliente_Id   = NE.EndCliente_Id " & vbCrLf & _
                    "			AND NF.EntradaSaida_Id = NE.EntradaSaida_Id " & vbCrLf & _
                    "			AND NF.Serie_Id        = NE.Serie_Id " & vbCrLf & _
                    "			AND NF.Nota_Id         = NE.Nota_Id " & vbCrLf & _
                    "	INNER JOIN Clientes C " & vbCrLf & _
                    "			ON C.Cliente_Id   = NF.Cliente_Id " & vbCrLf & _
                    "			AND C.Endereco_id = NF.EndCliente_Id " & vbCrLf & _
                    "	INNER JOIN SubOperacoes so " & vbCrLf & _
                    "			ON so.Operacao_Id      = NF.Operacao " & vbCrLf & _
                    "			AND so.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                    "WHERE NE.Encargo_Id = 'LIQUIDO' " & vbCrLf & _
                    "  AND NF.SITUACAO IN(1,4,7) " & vbCrLf & _
                    "  AND NF.Empresa_Id = '" & objPedido.CodigoEmpresa & "' " & vbCrLf & _
                    "  AND NF.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
                    "  AND NF.Pedido = " & objPedido.Codigo & vbCrLf & _
                    "GROUP BY NF.Movimento, NF.Empresa_id, NF.EndEmpresa_id, NF.Cliente_Id, NF.EndCliente_Id, C.Nome, " & vbCrLf & _
                    "		NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id, NF.Operacao, NF.SubOperacao " & vbCrLf & _
                    "ORDER BY NF.Movimento, NF.Nota_Id"
            '"  AND so.Financeiro = 'S' " & vbCrLf & _

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "ConsultaNotas")

            GridNotas.DataSource = ds.Tables(0)
            GridNotas.DataBind()

            Dim vlrNOTAS As Decimal = 0

            If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                If IsDBNull(ds.Tables(0).Compute("SUM(Valor)", "ES = 'S'")) Then
                    vlrNOTAS = ds.Tables(0).Compute("SUM(Valor)", "ES = 'E'")
                Else
                    vlrNOTAS = ds.Tables(0).Compute("SUM(Valor)", "ES = 'E'") - ds.Tables(0).Compute("SUM(Valor)", "ES = 'S'")
                End If
            Else
                If IsDBNull(ds.Tables(0).Compute("SUM(Valor)", "ES = 'E'")) Then
                    vlrNOTAS = ds.Tables(0).Compute("SUM(Valor)", "ES = 'S'")
                Else
                    vlrNOTAS = ds.Tables(0).Compute("SUM(Valor)", "ES = 'S'") - ds.Tables(0).Compute("SUM(Valor)", "ES = 'E'")
                End If
            End If

            txtTotalNotas.Text = vlrNOTAS.ToString("N2")

            txtCnpjDaEmpresa.Text = Funcoes.FormatarCpfCnpj(objPedido.CodigoEmpresa)
            txtNomeDaEmpresa.Text = objPedido.Empresa.Nome
            txtCidadeDaEmpresa.Text = objPedido.Empresa.Cidade & "/" & objPedido.Empresa.CodigoEstado

            txtCnpjCliente.Text = Funcoes.FormatarCpfCnpj(objPedido.CodigoCliente)
            txtNomeDoCliente.Text = objPedido.Cliente.Nome
            txtCidadeCliente.Text = objPedido.Cliente.Cidade & "/" & objPedido.Cliente.CodigoEstado

            For Each rowDep As [Lib].Negocio.PedidoXDeposito In objPedido.Depositos
                If rowDep.Tipo = "DE" Then
                    txtDeposito.Text = Funcoes.FormatarCpfCnpj(rowDep.Codigo)
                    txtNomeDoDeposito.Text = rowDep.Deposito.Nome
                    txtCidadeDeposito.Text = rowDep.Deposito.Cidade & "/" & rowDep.Deposito.CodigoEstado
                End If
            Next

            txtPedido.Text = objPedido.Codigo
            txtDataEntrega.Text = objPedido.DataEntregaInicial
            txtNaturezaDaOperacao.Text = objPedido.SubOperacao.Classe.ToString()
            txtOperacao.Text = objPedido.CodigoOperacao & "-" & objPedido.CodigoSubOperacao & " " & objPedido.SubOperacao.Descricao
            txtMoeda.Text = objPedido.Moeda.Descricao

            SalvarPedido()
            TabContainer1.ActiveTabIndex = 1

            If gridFinanceiro.Rows.Count = 1 AndAlso GridNotas.Rows.Count = 1 AndAlso CDec(txtTotalTitulos.Text) = CDec(txtTotalNotas.Text) Then
                lnkUmNovo.Parent.Visible = True
            Else
                TabParcelamento.Visible = True

                If objPedido.CodigoCondicaoPagamento > 0 Then ddlCondPagPed.SelectedValue = objPedido.CodigoCondicaoPagamento
                txtDataCondPagParcela.Text = objPedido.DataPedido

                Dim dtNotasXTitulos As New DataTable("listNotasXTitulos")
                dtNotasXTitulos.Columns.Add("ES", Type.GetType("System.String"))
                dtNotasXTitulos.Columns.Add("Serie", Type.GetType("System.String"))
                dtNotasXTitulos.Columns.Add("Nota", Type.GetType("System.String"))
                dtNotasXTitulos.Columns.Add("TipoFinanceiro", Type.GetType("System.String"))
                dtNotasXTitulos.Columns.Add("Provisao", Type.GetType("System.String"))
                dtNotasXTitulos.Columns.Add("Titulo", Type.GetType("System.String"))
                dtNotasXTitulos.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
                dtNotasXTitulos.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
                dtNotasXTitulos.Columns.Add("DeducaoOficial", Type.GetType("System.Decimal"))
                dtNotasXTitulos.Columns.Add("DescontoOficial", Type.GetType("System.Decimal"))
                dtNotasXTitulos.Columns.Add("AcrescimoOficial", Type.GetType("System.Decimal"))
                dtNotasXTitulos.Columns.Add("JurosOficial", Type.GetType("System.Decimal"))
                dtNotasXTitulos.Columns.Add("LiquidoOficial", Type.GetType("System.Decimal"))
                dtNotasXTitulos.Columns.Add("LiquidoMoeda", Type.GetType("System.Decimal"))

                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim nVencimento As String = String.Empty

                For Each rowNota As DataRow In ds.Tables(0).Rows
                    While rowNota("Valor") > 0
                        Dim drNxT As DataRow = dtNotasXTitulos.NewRow()
                        drNxT("ES") = rowNota("ES")
                        drNxT("Serie") = rowNota("Serie")
                        drNxT("Nota") = rowNota("Nota")

                        Dim nf As New [Lib].Negocio.NotaFiscal()
                        nf.CodigoEmpresa = rowNota("Empresa_Id")
                        nf.EnderecoEmpresa = rowNota("EndEmpresa_id")
                        nf.CodigoCliente = rowNota("Cliente")
                        nf.EnderecoCliente = rowNota("EndCliente")
                        nf.EntradaSaida = IIf(rowNota("ES") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                        nf.Codigo = rowNota("Nota")
                        nf.Serie = rowNota("Serie")
                        nf = New [Lib].Negocio.NotaFiscal(nf)

                        If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                            If rowNota("ES") = "E" Then
                                For Each rowTit As DataRow In dtTitulos.Rows
                                    If rowTit("TipoFinanceiro") = "P" AndAlso rowTit("DocumentoOficial") > 0 Then
                                        drNxT("TipoFinanceiro") = rowTit("TipoFinanceiro")
                                        drNxT("Provisao") = rowTit("Provisao")
                                        drNxT("Titulo") = rowTit("Codigo")

                                        If rowNota("Movimento") > CDate(rowTit("Vencimento")) Then
                                            drNxT("Vencimento") = rowNota("Movimento")
                                        Else
                                            drNxT("Vencimento") = rowTit("Vencimento")
                                        End If

                                        nVencimento = drNxT("Vencimento")

                                        If rowTit("DocumentoOficial") >= rowNota("Valor") Then
                                            drNxT("ValorOficial") = rowNota("Valor")
                                        Else
                                            drNxT("ValorOficial") = rowTit("DocumentoOficial")
                                        End If

                                        If vDeducoesP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("DeducaoOficial") = vDeducoesP
                                            Else
                                                drNxT("DeducaoOficial") = vDeducoesP / totalDeTitulosP
                                            End If
                                            vDeducoesP -= drNxT("DeducaoOficial")
                                        Else
                                            drNxT("DeducaoOficial") = 0
                                        End If

                                        If vDescontosP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("DescontoOficial") = vDescontosP
                                            Else
                                                drNxT("DescontoOficial") = vDescontosP / totalDeTitulosP
                                            End If
                                            vDescontosP -= drNxT("DescontoOficial")
                                        Else
                                            drNxT("DescontoOficial") = 0
                                        End If

                                        If vAcrescimosP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("AcrescimoOficial") = vAcrescimosP
                                            Else
                                                drNxT("AcrescimoOficial") = vAcrescimosP / totalDeTitulosP
                                            End If
                                            vAcrescimosP -= drNxT("AcrescimoOficial")
                                        Else
                                            drNxT("AcrescimoOficial") = 0
                                        End If

                                        If vJurosP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("JurosOficial") = vJurosP
                                            Else
                                                drNxT("JurosOficial") = vJurosP / totalDeTitulosP
                                            End If
                                            vJurosP -= drNxT("JurosOficial")
                                        Else
                                            drNxT("JurosOficial") = 0
                                        End If

                                        If rowTit("Provisao") = "1" Then i += 1

                                        drNxT("LiquidoOficial") = drNxT("ValorOficial") - drNxT("DeducaoOficial") - drNxT("DescontoOficial") + drNxT("AcrescimoOficial") + drNxT("JurosOficial")

                                        rowTit("DocumentoOficial") -= drNxT("ValorOficial")
                                        rowNota("Valor") -= drNxT("ValorOficial")

                                        If objPedido.CodigoMoeda = 3 Then
                                            drNxT("LiquidoMoeda") = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                                            If nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda) = 0 Then
                                                MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                                Exit Sub
                                            End If
                                        End If

                                        Exit For
                                    End If
                                Next
                            Else
                                For Each rowTit As DataRow In dtTitulos.Rows
                                    If rowTit("TipoFinanceiro") = "R" AndAlso rowTit("DocumentoOficial") > 0 Then
                                        drNxT("TipoFinanceiro") = rowTit("TipoFinanceiro")
                                        drNxT("Provisao") = rowTit("Provisao")
                                        drNxT("Titulo") = rowTit("Codigo")

                                        If rowNota("Movimento") > CDate(rowTit("Vencimento")) Then
                                            drNxT("Vencimento") = rowNota("Movimento")
                                        Else
                                            drNxT("Vencimento") = rowTit("Vencimento")
                                        End If

                                        nVencimento = drNxT("Vencimento")

                                        If rowTit("DocumentoOficial") >= rowNota("Valor") Then
                                            drNxT("ValorOficial") = rowNota("Valor")
                                        Else
                                            drNxT("ValorOficial") = rowTit("DocumentoOficial")
                                        End If

                                        If vDeducoesR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("DeducaoOficial") = vDeducoesR
                                            Else
                                                drNxT("DeducaoOficial") = vDeducoesR / totalDeTitulosR
                                            End If
                                            vDeducoesR -= drNxT("DeducaoOficial")
                                        Else
                                            drNxT("DeducaoOficial") = 0
                                        End If

                                        If vDescontosR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("DescontoOficial") = vDescontosR
                                            Else
                                                drNxT("DescontoOficial") = vDescontosR / totalDeTitulosR
                                            End If
                                            vDescontosR -= drNxT("DescontoOficial")
                                        Else
                                            drNxT("DescontoOficial") = 0
                                        End If

                                        If vAcrescimosR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("AcrescimoOficial") = vAcrescimosR
                                            Else
                                                drNxT("AcrescimoOficial") = vAcrescimosR / totalDeTitulosR
                                            End If
                                            vAcrescimosR -= drNxT("AcrescimoOficial")
                                        Else
                                            drNxT("AcrescimoOficial") = 0
                                        End If

                                        If vJurosR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("JurosOficial") = vJurosR
                                            Else
                                                drNxT("JurosOficial") = vJurosR / totalDeTitulosR
                                            End If
                                            vJurosR -= drNxT("JurosOficial")
                                        Else
                                            drNxT("JurosOficial") = 0
                                        End If

                                        If rowTit("Provisao") = "1" Then j += 1

                                        drNxT("LiquidoOficial") = drNxT("ValorOficial") - drNxT("DeducaoOficial") - drNxT("DescontoOficial") + drNxT("AcrescimoOficial") + drNxT("JurosOficial")

                                        rowTit("DocumentoOficial") -= drNxT("ValorOficial")
                                        rowNota("Valor") -= drNxT("ValorOficial")

                                        If objPedido.CodigoMoeda = 3 Then
                                            drNxT("LiquidoMoeda") = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                                            If nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda) = 0 Then
                                                MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                                Exit Sub
                                            End If
                                        End If

                                        Exit For
                                    End If
                                Next
                            End If
                        Else
                            If rowNota("ES") = "S" Then
                                For Each rowTit As DataRow In dtTitulos.Rows
                                    If rowTit("TipoFinanceiro") = "R" AndAlso rowTit("DocumentoOficial") > 0 Then
                                        drNxT("TipoFinanceiro") = rowTit("TipoFinanceiro")
                                        drNxT("Provisao") = rowTit("Provisao")
                                        drNxT("Titulo") = rowTit("Codigo")

                                        If rowNota("Movimento") > CDate(rowTit("Vencimento")) Then
                                            drNxT("Vencimento") = rowNota("Movimento")
                                        Else
                                            drNxT("Vencimento") = rowTit("Vencimento")
                                        End If

                                        nVencimento = drNxT("Vencimento")

                                        If rowTit("DocumentoOficial") >= rowNota("Valor") Then
                                            drNxT("ValorOficial") = rowNota("Valor")
                                        Else
                                            drNxT("ValorOficial") = rowTit("DocumentoOficial")
                                        End If

                                        If vDeducoesR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("DeducaoOficial") = vDeducoesR
                                            Else
                                                drNxT("DeducaoOficial") = vDeducoesR / totalDeTitulosR
                                            End If
                                            vDeducoesR -= drNxT("DeducaoOficial")
                                        Else
                                            drNxT("DeducaoOficial") = 0
                                        End If

                                        If vDescontosR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("DescontoOficial") = vDescontosR
                                            Else
                                                drNxT("DescontoOficial") = vDescontosR / totalDeTitulosR
                                            End If
                                            vDescontosR -= drNxT("DescontoOficial")
                                        Else
                                            drNxT("DescontoOficial") = 0
                                        End If

                                        If vAcrescimosR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("AcrescimoOficial") = vAcrescimosR
                                            Else
                                                drNxT("AcrescimoOficial") = vAcrescimosR / totalDeTitulosR
                                            End If
                                            vAcrescimosR -= drNxT("AcrescimoOficial")
                                        Else
                                            drNxT("AcrescimoOficial") = 0
                                        End If

                                        If vJurosR > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosR - 1 Then
                                                drNxT("JurosOficial") = vJurosR
                                            Else
                                                drNxT("JurosOficial") = vJurosR / totalDeTitulosR
                                            End If
                                            vJurosR -= drNxT("JurosOficial")
                                        Else
                                            drNxT("JurosOficial") = 0
                                        End If

                                        If rowTit("Provisao") = "1" Then j += 1

                                        drNxT("LiquidoOficial") = drNxT("ValorOficial") - drNxT("DeducaoOficial") - drNxT("DescontoOficial") + drNxT("AcrescimoOficial") + drNxT("JurosOficial")

                                        rowTit("DocumentoOficial") -= drNxT("ValorOficial")
                                        rowNota("Valor") -= drNxT("ValorOficial")

                                        If objPedido.CodigoMoeda = 3 Then
                                            drNxT("LiquidoMoeda") = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                                            If nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda) = 0 Then
                                                MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                                Exit Sub
                                            End If
                                        End If
                                        Exit For
                                    End If
                                Next
                            Else
                                For Each rowTit As DataRow In dtTitulos.Rows
                                    If rowTit("TipoFinanceiro") = "P" AndAlso rowTit("DocumentoOficial") > 0 Then
                                        drNxT("TipoFinanceiro") = rowTit("TipoFinanceiro")
                                        drNxT("Provisao") = rowTit("Provisao")
                                        drNxT("Titulo") = rowTit("Codigo")

                                        If rowNota("Movimento") > CDate(rowTit("Vencimento")) Then
                                            drNxT("Vencimento") = rowNota("Movimento")
                                        Else
                                            drNxT("Vencimento") = rowTit("Vencimento")
                                        End If

                                        nVencimento = drNxT("Vencimento")

                                        If rowTit("DocumentoOficial") >= rowNota("Valor") Then
                                            drNxT("ValorOficial") = rowNota("Valor")
                                        Else
                                            drNxT("ValorOficial") = rowTit("DocumentoOficial")
                                        End If

                                        If vDeducoesP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("DeducaoOficial") = vDeducoesP
                                            Else
                                                drNxT("DeducaoOficial") = vDeducoesP / totalDeTitulosP
                                            End If
                                            vDeducoesP -= drNxT("DeducaoOficial")
                                        Else
                                            drNxT("DeducaoOficial") = 0
                                        End If

                                        If vDescontosP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("DescontoOficial") = vDescontosP
                                            Else
                                                drNxT("DescontoOficial") = vDescontosP / totalDeTitulosP
                                            End If
                                            vDescontosP -= drNxT("DescontoOficial")
                                        Else
                                            drNxT("DescontoOficial") = 0
                                        End If

                                        If vAcrescimosP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("AcrescimoOficial") = vAcrescimosP
                                            Else
                                                drNxT("AcrescimoOficial") = vAcrescimosP / totalDeTitulosP
                                            End If
                                            vAcrescimosP -= drNxT("AcrescimoOficial")
                                        Else
                                            drNxT("AcrescimoOficial") = 0
                                        End If

                                        If vJurosP > 0 And rowTit("Provisao") = "1" Then
                                            If i = totalDeTitulosP - 1 Then
                                                drNxT("JurosOficial") = vJurosP
                                            Else
                                                drNxT("JurosOficial") = vJurosP / totalDeTitulosP
                                            End If
                                            vJurosP -= drNxT("JurosOficial")
                                        Else
                                            drNxT("JurosOficial") = 0
                                        End If

                                        If rowTit("Provisao") = "1" Then i += 1

                                        drNxT("LiquidoOficial") = drNxT("ValorOficial") - drNxT("DeducaoOficial") - drNxT("DescontoOficial") + drNxT("AcrescimoOficial") + drNxT("JurosOficial")

                                        rowTit("DocumentoOficial") -= drNxT("ValorOficial")
                                        rowNota("Valor") -= drNxT("ValorOficial")

                                        If objPedido.CodigoMoeda = 3 Then
                                            drNxT("LiquidoMoeda") = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                                            If nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda) = 0 Then
                                                MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                                Exit Sub
                                            End If
                                        End If

                                        Exit For
                                    End If
                                Next
                            End If
                        End If

                        If IsDBNull(drNxT("ValorOficial")) AndAlso rowNota("Valor") > 0 Then
                            If rowNota("ES") = "E" Then
                                'If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                                drNxT("TipoFinanceiro") = "P"
                                'Else
                                '    drNxT("TipoFinanceiro") = "R"
                                'End If
                            Else
                                'If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                                drNxT("TipoFinanceiro") = "R"
                                'Else
                                '    drNxT("TipoFinanceiro") = "P"
                                'End If
                            End If
                            drNxT("Provisao") = "2"
                            drNxT("Titulo") = "0"

                            If String.IsNullOrWhiteSpace(nVencimento) Then
                                drNxT("Vencimento") = objPedido.DataEntregaFinal
                            Else
                                drNxT("Vencimento") = nVencimento
                            End If

                            drNxT("ValorOficial") = rowNota("Valor")
                            rowNota("Valor") -= drNxT("ValorOficial")

                            drNxT("DeducaoOficial") = 0
                            drNxT("DescontoOficial") = 0
                            drNxT("AcrescimoOficial") = 0
                            drNxT("JurosOficial") = 0
                            drNxT("LiquidoOficial") = drNxT("ValorOficial") - drNxT("DeducaoOficial") - drNxT("DescontoOficial") + drNxT("AcrescimoOficial") + drNxT("JurosOficial")

                            If objPedido.CodigoMoeda = 3 Then
                                drNxT("LiquidoMoeda") = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                                If nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda) = 0 Then
                                    MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                    Exit Sub
                                End If
                            End If
                        End If

                        dtNotasXTitulos.Rows.Add(drNxT)

                    End While
                Next

                GridNotasXTitulos.DataSource = dtNotasXTitulos
                GridNotasXTitulos.DataBind()

                lnkVirtualizar.Parent.Visible = False
                lnkVincular.Parent.Visible = False

                If objPedido.Vencimentos.Any(Function(s) s.Movimento.Year < 2016 AndAlso s.Provisao = eProvisao.Baixa) Then
                    MsgBox(Me.Page, "Pedidos com Baixa anterior à 2016, Títulos serão apenas vinculados as Notas Fiscais e criados se necessário, não será refeita a contabilização.")
                    lnkVincular.Parent.Visible = True
                ElseIf objPedido.CodigoMoeda = 3 Then 'Pedidos ém dólar os Títulos serão apenas vinculados as Notas Fiscais e criados se necessário
                    lnkVincular.Parent.Visible = True
                Else
                    lnkVirtualizar.Parent.Visible = True
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkUmNovo_Click(sender As Object, e As EventArgs) Handles lnkUmNovo.Click
        Try
            RecuperarPedido()

            Dim Sqls As New ArrayList

            Dim parcela As New PedidoXParcela(objPedido)
            parcela.IUD = "I"
            parcela.CodigoParcela = 1
            parcela.DataVencimento = objPedido.Vencimentos(0).Vencimento
            parcela.Valor = objPedido.Vencimentos(0).ValorLiquidoOficial
            parcela.SalvarSql(Sqls)

            sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf & _
                  " values('" & GridNotas.Rows(0).Cells(1).Text & "'," & GridNotas.Rows(0).Cells(2).Text & ", " & vbCrLf & _
                  "'" & GridNotas.Rows(0).Cells(3).Text & "'," & GridNotas.Rows(0).Cells(4).Text & "," & vbCrLf & _
                  "'" & GridNotas.Rows(0).Cells(6).Text & "','" & GridNotas.Rows(0).Cells(7).Text & "'," & GridNotas.Rows(0).Cells(8).Text & "," & objPedido.Vencimentos(0).Codigo & ")"

            Sqls.Add(sql)

            If Banco.GravaBanco(Sqls) Then
                Limpar(False)
                lnkConsultar_Click(Nothing, Nothing)
            Else
                MsgBox(Me.Page, "Erro ao gravar registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdArquivoDeSaida_Click(sender As Object, e As EventArgs) Handles cmdArquivoDeSaida.Click
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & HiddenArquivo.Value)

            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(Server.MapPath("~/Files/" & HiddenArquivo.Value), FileMode.Open, FileAccess.Read)

                Do
                    If Response.IsClientConnected Then
                        length = download.Read(buffer, 0, bufferLength)
                        Response.OutputStream.Write(buffer, 0, length)
                        buffer = New Byte(bufferLength - 1) {}
                    Else
                        length = -1
                    End If
                Loop While length > 0
                Response.Flush()
                Response.End()
            Finally
                If download IsNot Nothing Then
                    download.Close()
                    cmdArquivoDeSaida.Visible = False
                End If
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkVirtualizar_Click(sender As Object, e As EventArgs) Handles lnkVirtualizar.Click
        If gridParcelas.Rows.Count = 0 Then
            MsgBox(Me.Page, "Parcelamento do Pedido não foi feito!")
            Exit Sub
        End If

        Dim processoFinalizado As Boolean = False

        Dim NomeArquivo As String
        Dim NomeArquivo2 As String = "Files/resumopedido" & txtPedido.Text & ".html"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/resumopedido" & txtPedido.Text & ".html"))

        Dim strHTMLraz As String = String.Empty
        Dim strHTMLadto As String = String.Empty
        Dim temBAdto As Boolean = False
        Dim seqAdto As Integer = 0
        Dim totalBaixaAdto As Decimal = 0
        Dim totalDebito As Decimal = 0
        Dim totalCredito As Decimal = 0

        Dim strHTML As String = "<html>" & vbCrLf & _
                                "   <head>" & vbCrLf & _
                                "     <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
                                "       <title>RESUMO DA CONVERSÃO DO PEDIDO</title>" & vbCrLf & _
                                "       <style type=""text/css"">" & vbCrLf & _
                                "           H6 {page-break-after:always}" & vbCrLf & _
                                "           A.A1 {font-family: Ms serif;font-size: 5pt;line height: 15px}" & vbCrLf & _
                                "           A.A2 {font-family: Ms serif;font-size: 6pt;line height: 15px}" & vbCrLf & _
                                "           A.A3 {font-family: Ms serif;font-size: 7pt;line height: 15px}" & vbCrLf & _
                                "           A.A4 {font-family: Ms serif;font-size: 8pt;line height: 15px}" & vbCrLf & _
                                "           A.A5 {font-family: Ms serif;font-size: 9pt;line height: 15px}" & vbCrLf & _
                                "           A.A6 {font-family: Ms serif;font-size: 10pt;line height: 15px}" & vbCrLf & _
                                "           A.A7 {font-family: Ms serif bold;font-size: 8pt;line height: 15px}" & vbCrLf & _
                                "           A.A8 {font-family: Ms serif bold;font-size: 9pt;line height: 15px}" & vbCrLf & _
                                "           A.A9 {font-family: Ms serif bold;font-size: 10pt;line height: 15px}" & vbCrLf & _
                                "           A.A10 {font-family: Ms serif bold;font-size: 11pt;line height: 15px}" & vbCrLf & _
                                "           A.A11 {font-family: Ms serif bold;font-size: 12pt;line height: 15px}" & vbCrLf & _
                                "           A.A12 {font-family: Ms serif bold;font-size: 13pt;line height: 15px}" & vbCrLf & _
                                "           A.A13 {font-family: Ms serif bold;font-size: 14pt;line height: 15px}" & vbCrLf & _
                                "       </style>" & vbCrLf & _
                                "   </head>" & vbCrLf & _
                                "   <bodyb text=#000000 bgcolor=#FFFFFF>"

        strHTML &= "<table border='0' width='100%'>" & vbCrLf & _
                   "<tr>" & vbCrLf & _
                   "<td>" & vbCrLf

        Try
            RecuperarPedido()

            If objPedido.Vencimentos.Count > 0 Then
                'Vencimentos Originais
                strHTML &= "<table border='0'>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='12'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='12'><A Class='A10'>VENCIMENTOS ORIGINAIS</A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td><A Class='A5'>PAG/REC</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>TITULO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>PROVISÃO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>VENCIMENTO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>BAIXA</A></td>" & vbCrLf & _
                           "<td align='center'><A Class='A5'>MOEDA</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>DEDUÇÃO</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>DESCONTO</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>ACRÉSCIMOS</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>JUROS</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>LÍQUIDO</A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf

                For Each parcela In objPedido.Vencimentos.ToList.OrderBy(Function(s) s.Provisao)
                    strHTML &= "<tr>" & vbCrLf & _
                              "<td><A Class='A5'>" & parcela.TipoFinanceiro & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & parcela.Codigo & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & parcela.Provisao.ToInt32 & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & parcela.DataProrrogacao & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & parcela.DataBaixa & "</A></td>" & vbCrLf
                    If parcela.CodigoMoeda = 1 Then
                        strHTML &= "<td align='center'><A Class='A5'>R$</A></td>" & vbCrLf
                    Else
                        strHTML &= "<td align='center'><A Class='A5'>U$</A></td>" & vbCrLf
                    End If

                    strHTML &= "<td align='right'><A Class='A5'>" & parcela.ValorDocumentoOficial & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & parcela.DeducoesOficial & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & parcela.DescontosOficial & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & parcela.AcrescimosOficial & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & parcela.JurosOficial & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & parcela.ValorLiquidoOficial & "</A></td>" & vbCrLf & _
                              "</tr>" & vbCrLf
                Next

                strHTML &= "<tr>" & vbCrLf & _
                           "<td colspan='12'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "</table>" & vbCrLf

                If objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa).Count > 0 Then
                    For Each tit In objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa)
                        Dim baixaAdto As New ListAdiantamentoBaixa(tit.Codigo)
                        If Not baixaAdto Is Nothing AndAlso baixaAdto.Count > 0 Then
                            'Baixas de Adiantamentos Originais
                            strHTML &= "<table border='0'>" & vbCrLf & _
                                       "<tr>" & vbCrLf & _
                                       "<td colspan='4'>" & vbCrLf & _
                                       "<hr />" & vbCrLf & _
                                       "</td>" & vbCrLf & _
                                       "</tr>" & vbCrLf & _
                                       "<tr>" & vbCrLf & _
                                       "<td colspan='4'><A Class='A10'>BAIXAS DE ADIANTAMENTOS ORIGINAIS</A></td>" & vbCrLf & _
                                       "</tr>" & vbCrLf & _
                                       "<tr>" & vbCrLf & _
                                       "<td><A Class='A5'>ADIANTAMENTO</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>SEQUÊNCIA</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>TÍTULO</A></td>" & vbCrLf & _
                                       "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                                       "</tr>" & vbCrLf

                            For Each baixa In baixaAdto
                                strHTML &= "<tr>" & vbCrLf & _
                                          "<td><A Class='A5'>" & baixa.Adiantamento.Codigo & "</A></td>" & vbCrLf & _
                                          "<td><A Class='A5'>" & baixa.Sequencia & "</A></td>" & vbCrLf & _
                                          "<td><A Class='A5'>" & baixa.CodigoTitulo & "</A></td>" & vbCrLf & _
                                          "<td align='right'><A Class='A5'>" & baixa.ValorOficial & "</A></td>" & vbCrLf & _
                                          "</tr>" & vbCrLf
                            Next

                            strHTML &= "<tr>" & vbCrLf & _
                                       "<td colspan='4'>" & vbCrLf & _
                                       "<hr />" & vbCrLf & _
                                       "</td>" & vbCrLf & _
                                       "</tr>" & vbCrLf & _
                                       "</table>" & vbCrLf
                        End If
                    Next

                    'Contabilização Original
                    strHTML &= "<table border='0'>" & vbCrLf & _
                               "<tr>" & vbCrLf & _
                               "<td colspan='10'>" & vbCrLf & _
                               "<hr />" & vbCrLf & _
                               "</td>" & vbCrLf & _
                               "</tr>" & vbCrLf & _
                               "<tr>" & vbCrLf & _
                               "<td colspan='10'><A Class='A10'>CONTABILIZAÇÃO ORIGINAL</A></td>" & vbCrLf & _
                               "</tr>" & vbCrLf & _
                               "<tr>" & vbCrLf & _
                               "<td><A Class='A5'>EMPRESA</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>CONTA</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>CLIENTE</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>MOVIMENTO</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>TÍTULO</A></td>" & vbCrLf & _
                               "<td align='right'><A Class='A5'>DÉBITO OFICIAL</A></td>" & vbCrLf & _
                               "<td align='right'><A Class='A5'>CRÉDITO OFICIAL</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>HISTÓRICO</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>USUÁRIO</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>DATA</A></td>" & vbCrLf & _
                               "</tr>" & vbCrLf

                    For Each tit In objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa)
                        Dim titulo As New Titulo(tit.Codigo, tit.TipoFinanceiro)
                        Dim razao As New ListRazao(titulo)

                        For Each raz In razao
                            strHTML &= "<tr>" & vbCrLf & _
                                       "<td><A Class='A5'>" & raz.CodigoEmpresa & "-" & raz.EndEmpresa & "</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>" & raz.CodigoConta & "</A></td>" & vbCrLf

                            If String.IsNullOrWhiteSpace(raz.CodigoCliente) Then
                                strHTML &= "<td><A Class='A5'></A></td>" & vbCrLf
                            Else
                                strHTML &= "<td><A Class='A5'>" & raz.CodigoCliente & "-" & raz.EnderecoCliente & "</A></td>" & vbCrLf
                            End If

                            strHTML &= "<td><A Class='A5'>" & raz.Movimento.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>" & raz.CodigoTitulo & "</A></td>" & vbCrLf & _
                                       "<td align='right'><A Class='A5'>" & raz.DebitoOficial & "</A></td>" & vbCrLf & _
                                       "<td align='right'><A Class='A5'>" & raz.CreditoOficial & "</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>" & raz.Historico & "</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>" & raz.UsuarioInclusao & "</A></td>" & vbCrLf & _
                                       "<td><A Class='A5'>" & raz.UsuarioDataInclusao.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                                       "</tr>" & vbCrLf

                            If raz.DebitoOficial > 0 Then
                                totalDebito += raz.DebitoOficial
                            Else
                                totalCredito += raz.CreditoOficial
                            End If
                        Next
                    Next

                    strHTML &= "<tr>" & vbCrLf & _
                               "<td colspan='10'>" & vbCrLf & _
                               "<hr />" & vbCrLf & _
                               "</td>" & vbCrLf & _
                               "</tr>" & vbCrLf & _
                               "<tr>" & vbCrLf & _
                               "<td colspan='5'><A Class='A5'><B>TOTAL ->" & vbCrLf & _
                               "</B></A></td>" & vbCrLf & _
                               "<td align='right'><A Class='A5'><B>" & totalDebito & vbCrLf & _
                               "</B></A></td>" & vbCrLf & _
                               "<td align='right'><A Class='A5'><B>" & totalCredito & vbCrLf & _
                               "</B></A></td>" & vbCrLf & _
                               "<td colspan='3'>" & vbCrLf & _
                               "</td>" & vbCrLf & _
                               "</tr>" & vbCrLf & _
                               "<tr>" & vbCrLf & _
                               "<td colspan='10'>" & vbCrLf & _
                               "<hr />" & vbCrLf & _
                               "</td>" & vbCrLf & _
                               "</tr>" & vbCrLf & _
                               "</table>" & vbCrLf
                End If
            End If

            Dim Sqls As New ArrayList
            totalDebito = 0
            totalCredito = 0

            'Cria parcelas no PedidoXParcela
            strHTML &= "<table border='0'>" & vbCrLf & _
                       "<tr>" & vbCrLf & _
                       "<td colspan='3'>" & vbCrLf & _
                       "<hr />" & vbCrLf & _
                       "</td>" & vbCrLf & _
                       "</tr>" & vbCrLf & _
                       "<tr>" & vbCrLf & _
                       "<td colspan='3'><A Class='A10'>PEDIDOxPARCELA</A></td>" & vbCrLf & _
                       "</tr>" & vbCrLf & _
                       "<tr>" & vbCrLf & _
                       "<td><A Class='A5'>PARCELA</A></td>" & vbCrLf & _
                       "<td><A Class='A5'>VENCIMENTO</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                       "</tr>" & vbCrLf

            For Each pRow As GridViewRow In gridParcelas.Rows
                Dim parcela As New PedidoXParcela(objPedido)
                parcela.IUD = "I"
                parcela.CodigoParcela = pRow.Cells(1).Text
                parcela.DataVencimento = pRow.Cells(2).Text
                parcela.Valor = CDec(pRow.Cells(3).Text)
                parcela.SalvarSql(Sqls)

                strHTML &= "<tr>" & vbCrLf & _
                          "<td><A Class='A5'>" & parcela.CodigoParcela & "</A></td>" & vbCrLf & _
                          "<td><A Class='A5'>" & parcela.DataVencimento & "</A></td>" & vbCrLf & _
                          "<td align='right'><A Class='A5'>" & parcela.Valor & "</A></td>" & vbCrLf & _
                          "</tr>" & vbCrLf
            Next

            strHTML &= "<tr>" & vbCrLf & _
                       "<td colspan='3'>" & vbCrLf & _
                       "<hr />" & vbCrLf & _
                       "</td>" & vbCrLf & _
                       "</tr>" & vbCrLf & _
                       "</table>" & vbCrLf


            'Vencimentos Novos
            strHTML &= "<table border='0'>" & vbCrLf & _
                       "<tr>" & vbCrLf & _
                       "<td colspan='12'>" & vbCrLf & _
                       "<hr />" & vbCrLf & _
                       "</td>" & vbCrLf & _
                       "</tr>" & vbCrLf & _
                       "<tr>" & vbCrLf & _
                       "<td colspan='12'><A Class='A10'>VENCIMENTOS NOVOS</A></td>" & vbCrLf & _
                       "</tr>" & vbCrLf & _
                       "<tr>" & vbCrLf & _
                       "<td align='center'><A Class='A5'>PAG/REC</A></td>" & vbCrLf & _
                       "<td><A Class='A5'>TITULO</A></td>" & vbCrLf & _
                       "<td align='center'><A Class='A5'>PROVISÃO</A></td>" & vbCrLf & _
                       "<td><A Class='A5'>VENCIMENTO</A></td>" & vbCrLf & _
                       "<td><A Class='A5'>BAIXA</A></td>" & vbCrLf & _
                       "<td align='center'><A Class='A5'>MOEDA</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>DEDUÇÃO</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>DESCONTO</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>ACRÉSCIMOS</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>JUROS</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>LÍQUIDO</A></td>" & vbCrLf & _
                       "</tr>" & vbCrLf

            For Each nRow As GridViewRow In GridNotasXTitulos.Rows
                'Dim dTitulo As String = nRow.Cells(6).Text

                Dim nf As New [Lib].Negocio.NotaFiscal()

                For Each rowNF As GridViewRow In GridNotas.Rows
                    If nRow.Cells(0).Text = rowNF.Cells(6).Text AndAlso nRow.Cells(1).Text = rowNF.Cells(7).Text AndAlso nRow.Cells(2).Text = rowNF.Cells(8).Text Then
                        nf.CodigoEmpresa = rowNF.Cells(1).Text
                        nf.EnderecoEmpresa = rowNF.Cells(2).Text
                        nf.CodigoCliente = rowNF.Cells(3).Text
                        nf.EnderecoCliente = rowNF.Cells(4).Text
                        nf.EntradaSaida = IIf(rowNF.Cells(6).Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                        nf.Codigo = rowNF.Cells(8).Text
                        nf.Serie = rowNF.Cells(7).Text
                        nf = New [Lib].Negocio.NotaFiscal(nf)
                        Exit For
                    End If
                Next

                If CInt(nRow.Cells(5).Text) = "0" Then
                    If Not nf.Itens Is Nothing AndAlso nf.Itens.Count > 0 Then
                        Dim Tit As New Titulo

                        'NUMERADOR DOS TITULOS
                        Dim SqlN As String = "exec sp_Numerador '" & HttpContext.Current.Session("ssNomeServidor") & "',0," & [Lib].Negocio.eTiposNumerador.Titulo & ""
                        Dim dsN As New DataSet
                        dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                        Dim CodigoNumerador As Integer = 0
                        If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                            CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                        End If

                        If Not CodigoNumerador > 0 Then
                            MsgBox(Me.Page, "Numerador de Títulos não cadastrado!")
                            Exit Sub
                        End If

                        Tit.IndiceFixo = False

                        Tit.IUD = "I"
                        Tit.Codigo = CodigoNumerador
                        Tit.Sequencia = 0

                        'Caso Pedido AFixar vire PROVISAO = 4 (AFIXAR) - FURLAN - 04/10/2016
                        If objPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
                            Tit.CodigoProvisao = 4
                        Else
                            Tit.CodigoProvisao = 2
                        End If

                        If nf.EntradaSaida = eEntradaSaida.Entrada Then
                            Tit.ReceberPagar = "P"
                            If nf.SubOperacao.Devolucao Then
                                Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                                Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraVenda
                            Else
                                Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                                Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraCompra
                            End If
                        Else
                            Tit.ReceberPagar = "R"
                            If nf.SubOperacao.Devolucao Then
                                Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                                Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraCompra
                            Else
                                Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                                Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraVenda
                            End If
                        End If

                        Tit.Tributo = ""

                        Tit.CodigoEmpresaPedido = nf.Pedido.CodigoEmpresa
                        Tit.EndEmpresaPedido = nf.Pedido.EnderecoEmpresa
                        Tit.CodigoPedido = nf.CodigoPedido
                        Tit.CodigoPedidoFixacao = 0
                        Tit.CodigoProcuracao = nf.CodigoProcuracao
                        Tit.DataMoeda = CDate(nRow.Cells(6).Text) ' nf.Movimento.ToString("dd-MM-yyyy")
                        Tit.CodigoIndexador = nf.Pedido.CodigoIndexador
                        Tit.CodigoMoeda = nf.Pedido.CodigoMoeda
                        Tit.Movimento = nf.Movimento
                        Tit.Vencimento = Funcoes.ValidaDataUtil(nf.Empresa.Codigo, nf.Empresa.CodigoEndereco, CDate(nRow.Cells(6).Text))

                        Tit.Prorrogacao = Funcoes.ValidaDataUtil(nf.Empresa.Codigo, nf.Empresa.CodigoEndereco, CDate(nRow.Cells(6).Text))
                        Tit.Baixa = Funcoes.ValidaDataUtil(nf.Empresa.Codigo, nf.Empresa.CodigoEndereco, CDate(nRow.Cells(6).Text))
                        Tit.CodigoTipoPgto = 1
                        Tit.CodigoSituacao = 1
                        Tit.Lote = 70
                        Tit.CodigoUnidadeDeNegocio = nf.Pedido.CodigoUnidadeNegocio
                        Tit.CodigoEmpresa = nf.CodigoEmpresa
                        Tit.EnderecoEmpresa = nf.EnderecoEmpresa
                        Tit.CodigoCliente = nf.Pedido.CodigoCliente
                        Tit.EndCliente = nf.Pedido.EnderecoCliente
                        Tit.CodigoBancoCliente = 0
                        Tit.CodigoAgenciaCliente = ""
                        Tit.DigitoAgenciaCliente = ""
                        Tit.ContaCliente = ""
                        Tit.DigitoContaCliente = ""
                        Tit.Cheque = False
                        Tit.Slips = False
                        Tit.Recibo = False
                        Tit.Aviso = False
                        Tit.ReciboDeposito = False

                        Tit.ValorDoDocumento = CDec(nRow.Cells(7).Text)

                        If nf.Codigo > 0 AndAlso nf.Pedido.CodigoMoeda = 3 Then
                            Tit.MoedaValorDoDocumento = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                            If Tit.MoedaValorDoDocumento = 0 Then
                                MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                Exit Sub
                            End If
                        End If

                        Tit.Historico = "REF. NF " & nf.Codigo & "-" & nf.Serie & "-" & IIf(nf.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Pedido: " & nf.CodigoPedido & " / " & nf.Cliente.Nome
                        Tit.CodigoDeBarras = ""
                        Tit.CodigoDigitado = False
                        Tit.CodigoDeBarrasPreImpresso = False

                        Tit.CodigoDestinatario = nf.Pedido.CodigoCliente
                        Tit.EndDestinatario = nf.Pedido.EnderecoCliente
                        Tit.NomeDoDestinatario = ""
                        Tit.Destinacao = ""
                        Tit.Solicitacao = 0

                        Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
                        Tit.RegistroMestre = 0
                        Tit.Observacoes = ""
                        Tit.SituacaoBancaria = 0
                        Tit.NumeroDoCheque = 0
                        Tit.CodigoAdiantamento = 0
                        Tit.TaxaAdto = 0

                        Tit.UsuarioInclusao = nf.UsuarioInclusao
                        Tit.UsuarioInclusaoData = nf.DataInclusao

                        Tit.SalvarSql(Sqls)

                        strHTML &= "<tr>" & vbCrLf & _
                                  "<td align='center'><A Class='A5'>" & Tit.ReceberPagar & "</A></td>" & vbCrLf & _
                                  "<td><A Class='A5'>" & Tit.Codigo & "</A></td>" & vbCrLf & _
                                  "<td align='center'><A Class='A5'>" & Tit.CodigoProvisao & "</A></td>" & vbCrLf & _
                                  "<td><A Class='A5'>" & Tit.Prorrogacao.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                                  "<td><A Class='A5'>" & Tit.Baixa.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf
                        If Tit.CodigoMoeda = 1 Then
                            strHTML &= "<td align='center'><A Class='A5'>R$</A></td>" & vbCrLf
                        Else
                            strHTML &= "<td align='center'><A Class='A5'>U$</A></td>" & vbCrLf
                        End If

                        strHTML &= "<td align='right'><A Class='A5'>" & Tit.ValorDoDocumento & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & Tit.Deducoes & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & Tit.Descontos & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & Tit.Acrescimos & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & Tit.Juros & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & Tit.ValorLiquido & "</A></td>" & vbCrLf & _
                                  "</tr>" & vbCrLf
                    Else
                        MsgBox(Me.Page, "NF " & nf.Codigo & " com problema, entre em contato com o Suporte.")
                        Exit Sub
                    End If
                Else
                    If Not nf.Itens Is Nothing AndAlso nf.Itens.Count > 0 Then
                        Dim titulo As New Titulo(nRow.Cells(5).Text, nRow.Cells(3).Text)
                        Dim baixaAdto As New ListAdiantamentoBaixa(titulo.Codigo)

                        'NUMERADOR DOS TITULOS
                        Dim SqlN As String = "exec sp_Numerador '" & HttpContext.Current.Session("ssNomeServidor") & "',0," & [Lib].Negocio.eTiposNumerador.Titulo & ""
                        Dim dsN As New DataSet
                        dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                        Dim CodigoNumerador As Integer = 0
                        If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                            CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                        End If

                        If Not CodigoNumerador > 0 Then
                            MsgBox(Me.Page, "Numerador de Títulos não cadastrado!")
                            Exit Sub
                        End If

                        'GRAVAR NOVO TÍTULO
                        titulo.IUD = "I"
                        titulo.Codigo = CodigoNumerador
                        titulo.Movimento = nf.Movimento
                        titulo.Vencimento = CDate(nRow.Cells(6).Text)
                        titulo.Prorrogacao = CDate(nRow.Cells(6).Text)
                        titulo.ValorDoDocumento = CDec(nRow.Cells(7).Text)
                        titulo.Deducoes = CDec(nRow.Cells(8).Text)
                        titulo.Descontos = CDec(nRow.Cells(9).Text)
                        titulo.Acrescimos = CDec(nRow.Cells(10).Text)
                        titulo.Juros = CDec(nRow.Cells(11).Text)

                        If nf.Codigo > 0 AndAlso nf.Pedido.CodigoMoeda = 3 Then
                            titulo.MoedaValorDoDocumento = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                            If titulo.MoedaValorDoDocumento = 0 Then
                                MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                                Exit Sub
                            End If
                        End If

                        'Caso seja Provisão deve virar Previsão - Furlan - 04/10/2016
                        If titulo.CodigoProvisao = 3 Then titulo.CodigoProvisao = 2

                        'Caso Nota Fiscal AFixar ou Complementação mais não tenha Código da Fixacão vire PROVISAO = 4 (AFIXAR) - FURLAN - 04/10/2016
                        If objPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR AndAlso Not titulo.CodigoProvisao = 1 Then titulo.CodigoProvisao = 4

                        'If titulo.CodigoMoeda = 1 Then
                        '    If titulo.CodigoIndexador = 99 Then
                        '        titulo.MoedaValorDoDocumento = 0
                        '    Else
                        '        titulo.MoedaValorDoDocumento = Funcoes.ConverteParaMoedaExtrangeira(titulo.ValorDoDocumento, titulo.CodigoIndexador, titulo.Baixa)
                        '    End If
                        'End If

                        strHTML &= "<tr>" & vbCrLf & _
                                  "<td align='center'><A Class='A5'>" & titulo.ReceberPagar & "</A></td>" & vbCrLf & _
                                  "<td><A Class='A5'>" & titulo.Codigo & "</A></td>" & vbCrLf & _
                                  "<td align='center'><A Class='A5'>" & titulo.CodigoProvisao & "</A></td>" & vbCrLf & _
                                  "<td><A Class='A5'>" & titulo.Prorrogacao.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                                  "<td><A Class='A5'>" & titulo.Baixa.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf
                        If titulo.CodigoMoeda = 1 Then
                            strHTML &= "<td align='center'><A Class='A5'>R$</A></td>" & vbCrLf
                        Else
                            strHTML &= "<td align='center'><A Class='A5'>U$</A></td>" & vbCrLf
                        End If

                        strHTML &= "<td align='right'><A Class='A5'>" & titulo.ValorDoDocumento & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & titulo.Deducoes & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & titulo.Descontos & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & titulo.Acrescimos & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & titulo.Juros & "</A></td>" & vbCrLf & _
                                  "<td align='right'><A Class='A5'>" & titulo.ValorLiquido & "</A></td>" & vbCrLf & _
                                  "</tr>" & vbCrLf

                        titulo.SalvarSql(Sqls, False)

                        'GRAVAR RAZÃO CASO TENHA
                        If titulo.CodigoProvisao = 1 Then Contabilizar(Sqls, titulo, strHTMLraz, totalDebito, totalCredito)

                        'GRAVAR BAIXA DE ADIANTAMENTO
                        If Not baixaAdto Is Nothing AndAlso baixaAdto.Count > 0 Then
                            temBAdto = True
                            If seqAdto = 0 Then seqAdto = baixaAdto(0).Sequencia
                            For Each baixa In baixaAdto
                                baixa.IUD = "I"
                                seqAdto += 1
                                baixa.Sequencia = seqAdto
                                baixa.CodigoTitulo = CodigoNumerador

                                If titulo.ReceberPagar = "P" Then
                                    If titulo.Carteira.BaixaAdiantamento Then
                                        baixa.ValorOficial = titulo.ValorDoDocumento
                                    Else
                                        baixa.ValorOficial = titulo.ValorLiquido
                                    End If
                                Else
                                    baixa.ValorOficial = titulo.ValorDoDocumento
                                End If

                                baixa.ValorMoeda = 0
                                baixa.VariacaoOficial = 0
                                baixa.SalvarSql(Sqls)

                                'Baixas de Adiantamentos Novos
                                strHTMLadto &= "<tr>" & vbCrLf & _
                                          "<td><A Class='A5'>" & baixa.Adiantamento.Codigo & "</A></td>" & vbCrLf & _
                                          "<td><A Class='A5'>" & baixa.Sequencia & "</A></td>" & vbCrLf & _
                                          "<td><A Class='A5'>" & baixa.CodigoTitulo & "</A></td>" & vbCrLf & _
                                          "<td align='right'><A Class='A5'>" & baixa.ValorOficial & "</A></td>" & vbCrLf & _
                                          "</tr>" & vbCrLf

                                totalBaixaAdto += baixa.ValorOficial
                            Next
                        End If

                        'RELACIONAR NOTA FISCAL COM NOVO TÍTULO
                        sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf & _
                              " values('" & titulo.CodigoEmpresaPedido & "'," & titulo.EndEmpresaPedido & ", " & vbCrLf & _
                              "'" & titulo.CodigoCliente & "'," & titulo.EndCliente & "," & vbCrLf & _
                              "'" & nRow.Cells(0).Text & "','" & nRow.Cells(1).Text & "'," & nRow.Cells(2).Text & "," & titulo.Codigo & ")"

                        Sqls.Add(sql)
                    Else
                        MsgBox(Me.Page, "NF " & nf.Codigo & " com problema, entre em contato com o Suporte.")
                        Exit Sub
                    End If
                End If
            Next

            strHTML &= "<tr>" & vbCrLf & _
                       "<td colspan='12'>" & vbCrLf & _
                       "<hr />" & vbCrLf & _
                       "</td>" & vbCrLf & _
                       "</tr>" & vbCrLf & _
                       "</table>" & vbCrLf

            If temBAdto Then
                strHTML &= "<table border='0'>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='4'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='4'><A Class='A10'>BAIXA DE ADIANTAMENTO NOVO</A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td><A Class='A5'>ADIANTAMENTO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>SEQUÊNCIA</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>TÍTULO</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf

                strHTML &= strHTMLadto

                strHTML &= "<tr>" & vbCrLf & _
                           "<td colspan='4'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='3'><A Class='A5'><B>TOTAL -></B></A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'><B>" & totalBaixaAdto & "</B></A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='4'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "</table>" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(strHTMLraz) Then
                'Contabilização Nova
                strHTML &= "<table border='0'>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='10'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='10'><A Class='A10'>CONTABILIZAÇÃO NOVA</A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td><A Class='A5'>EMPRESA</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>CONTA</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>CLIENTE</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>MOVIMENTO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>TÍTULO</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>DÉBITO OFICIAL</A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'>CRÉDITO OFICIAL</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>HISTÓRICO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>USUÁRIO</A></td>" & vbCrLf & _
                           "<td><A Class='A5'>DATA</A></td>" & vbCrLf & _
                           "</tr>" & vbCrLf

                strHTML &= strHTMLraz

                strHTML &= "<tr>" & vbCrLf & _
                           "<td colspan='10'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='5'><A Class='A5'><B>TOTAL -></B></A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'><B>" & totalDebito & vbCrLf & _
                           "</B></A></td>" & vbCrLf & _
                           "<td align='right'><A Class='A5'><B>" & totalCredito & vbCrLf & _
                           "</B></A></td>" & vbCrLf & _
                           "<td colspan='3'>" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "<tr>" & vbCrLf & _
                           "<td colspan='10'>" & vbCrLf & _
                           "<hr />" & vbCrLf & _
                           "</td>" & vbCrLf & _
                           "</tr>" & vbCrLf & _
                           "</table>" & vbCrLf
            End If

            'Cancela os Titulos anteriores e deleta razão
            For Each row As GridViewRow In gridFinanceiro.Rows
                If row.Cells(0).Text = "P" Then
                    sql = "Update ContasAPagar set Situacao = 3 where Registro_id = " & row.Cells(1).Text
                Else
                    sql = "Update ContasAReceber set Situacao = 3 where Registro_id = " & row.Cells(1).Text
                End If
                Sqls.Add(sql)

                sql = "Delete From Razao where Titulo = " & row.Cells(1).Text
                Sqls.Add(sql)

                sql = "DELETE AdiantamentosXBaixas WHERE Titulo = " & row.Cells(1).Text
                Sqls.Add(sql)
            Next

            'Atualiza Momento Financeiro do Pedido 
            sql = "Update Pedidos set MomentoFinanceiro = 9, CondicaoPagamento = " & ddlCondPagPed.SelectedValue & vbCrLf & _
                "where Empresa_id = '" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                "  and EndEmpresa_id = " & objPedido.EnderecoEmpresa & vbCrLf & _
                "  and Pedido_id = " & objPedido.Codigo
            Sqls.Add(sql)

            If Banco.GravaBanco(Sqls) Then
                processoFinalizado = True
                HiddenArquivo.Value = "resumopedido" & txtPedido.Text & ".html"
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            strHTML &= "</td>" & vbCrLf & _
                      "</tr>" & vbCrLf & _
                      "</table>" & vbCrLf & _
                      "</body>" & vbCrLf & _
                      "</html>"

            strm.WriteLine(strHTML)

            strm.Close()
            strm.Dispose()

            If processoFinalizado Then
                Limpar(False)
                lnkConsultar_Click(Nothing, Nothing)

                cmdArquivoDeSaida.Visible = True

                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
            End If
        End Try
    End Sub

    Private Sub Contabilizar(ByRef Sqls As ArrayList, ByRef tit As Titulo, ByRef strHTMLraz As String, ByRef totalDebito As Decimal, ByRef totalCredito As Decimal)
        '------------------------------------------
        ' Gravar Valor do Documento
        '------------------------------------------
        Dim CodigoConta As String = String.Empty
        Dim temCliente As Boolean = False
        Dim Historico As String = String.Empty
        Dim DEBCRE As String = String.Empty

        Dim Carteira As New CarteiraFinanceira(tit.CodigoCarteira)
        CodigoConta = Carteira.CodigoContaCliente

        Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Carteira.CodigoContaCliente) 'Conta sem tributo
        temCliente = objPlaConta.TemCliente

        If Not String.IsNullOrWhiteSpace(tit.Tributo) Then                                    'Conta com tributo
            Dim Encargo As New Encargo(tit.Tributo)
            CodigoConta = Encargo.ContaDebito
            Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, CodigoConta)
            temCliente = objPlaContaTributo.TemCliente
        End If

        If tit.ReceberPagar = "P" Then
            DEBCRE = "D"
        Else
            DEBCRE = "C"
        End If

        LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, tit.CodigoEmpresa, tit.EnderecoEmpresa, CodigoConta, tit.CodigoCliente, tit.EndCliente, temCliente, DEBCRE, tit.ValorDoDocumento, tit.MoedaValorDoDocumento)

        temCliente = False

        '-------Descontos-----------
        If tit.Descontos > 0 Then
            CodigoConta = Carteira.CodigoContaDesconto 'Conta de Descontos Obtidos

            Dim objPlaContaDesconto As New [Lib].Negocio.PlanoDeConta("", 0, CodigoConta)
            temCliente = objPlaContaDesconto.TemCliente

            If tit.ReceberPagar = "P" Then
                DEBCRE = "C"
            Else
                DEBCRE = "D"
            End If

            LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, tit.CodigoEmpresa, tit.EnderecoEmpresa, CodigoConta, tit.CodigoCliente, tit.EndCliente, temCliente, DEBCRE, tit.Descontos, tit.MoedaDescontos)
        End If

        temCliente = False

        '----Deducoes--------------
        If tit.Deducoes > 0 Then
            CodigoConta = Carteira.CodigoContaDeducao 'Conta de Deduções

            Dim objPlaContaDeducoes As New [Lib].Negocio.PlanoDeConta("", 0, CodigoConta)
            temCliente = objPlaContaDeducoes.TemCliente

            If tit.ReceberPagar = "P" Then
                DEBCRE = "C"
            Else
                DEBCRE = "D"
            End If

            LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, tit.CodigoEmpresa, tit.EnderecoEmpresa, CodigoConta, tit.CodigoCliente, tit.EndCliente, temCliente, DEBCRE, tit.Deducoes, tit.MoedaDeducoes)
        End If

        temCliente = False

        '----Juros--------------
        If tit.Juros > 0 Then
            CodigoConta = Carteira.CodigoContaJuro 'Conta de Juros

            Dim objPlaContaJuros As New [Lib].Negocio.PlanoDeConta("", 0, CodigoConta)
            temCliente = objPlaContaJuros.TemCliente

            If tit.ReceberPagar = "P" Then
                DEBCRE = "D"
            Else
                DEBCRE = "C"
            End If

            LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, tit.CodigoEmpresa, tit.EnderecoEmpresa, CodigoConta, tit.CodigoCliente, tit.EndCliente, temCliente, DEBCRE, tit.Juros, tit.MoedaJuros)
        End If

        temCliente = False

        '----Acrescimos--------------
        If tit.Acrescimos > 0 Then
            CodigoConta = Carteira.CodigoContaAcrescimo 'Conta de Acréscimos

            Dim objPlaContaAcrescimo As New [Lib].Negocio.PlanoDeConta("", 0, CodigoConta)
            temCliente = objPlaContaAcrescimo.TemCliente

            If tit.ReceberPagar = "P" Then
                DEBCRE = "D"
            Else
                DEBCRE = "C"
            End If

            LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, tit.CodigoEmpresa, tit.EnderecoEmpresa, CodigoConta, tit.CodigoCliente, tit.EndCliente, temCliente, DEBCRE, tit.Acrescimos, tit.MoedaAcrescimos)
        End If

        temCliente = False

        '-----------------------
        'Gravar Líquido
        '-----------------------
        If String.IsNullOrWhiteSpace(tit.CarteiraAdto) Then
            CodigoConta = tit.ContaContabilPagadora
        Else
            Dim CarteiraAdto As New CarteiraFinanceira(tit.CarteiraAdto)
            CodigoConta = CarteiraAdto.CodigoContaCliente
        End If

        Dim objPlaContaLiquido As New [Lib].Negocio.PlanoDeConta("", 0, CodigoConta)
        temCliente = objPlaContaLiquido.TemCliente

        If tit.ReceberPagar = "P" Then
            DEBCRE = "C"
        Else
            DEBCRE = "D"
        End If

        LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, tit.CodigoEmpresaPagadora, tit.EndEmpresaPagadora, CodigoConta, tit.CodigoCliente, tit.EndCliente, temCliente, DEBCRE, tit.ValorLiquido, tit.MoedaValorLiquido)

        temCliente = False

        '-------------------------------------------
        'Transferencias Financeiras
        '-------------------------------------------
        sql = " SELECT EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, " & vbCrLf & _
              "        ContaContabil, ClienteContabil,EndClienteContabil,DebitoCredito " & vbCrLf & _
              "   FROM TransferenciasFinanceiras " & vbCrLf & _
              "  WHERE EmpresaDebito   ='" & tit.CodigoEmpresa & "'" & vbCrLf & _
              "    and EnderecoDebito  = " & tit.EnderecoEmpresa & vbCrLf & _
              "    and EmpresaCredito  ='" & tit.CodigoEmpresaPagadora & "'" & vbCrLf & _
              "    and EnderecoCredito = " & tit.EndEmpresaPagadora

        For Each DrT As DataRow In Banco.ConsultaDataSet(sql, "Transferencias").Tables(0).Rows
            LanctosContabeis(totalDebito, totalCredito, strHTMLraz, Sqls, tit, DrT("EmpresaContabil"), DrT("EnderecoContabil"), DrT("ContaContabil"), DrT("ClienteContabil"), DrT("EndClienteContabil"), temCliente, DrT("DebitoCredito"), tit.ValorLiquido, tit.MoedaValorLiquido)
        Next
    End Sub

    Private Sub LanctosContabeis(ByRef totalDebito As Decimal, ByRef totalCredito As Decimal, ByRef strHTMLraz As String, ByRef Sqls As ArrayList, ByVal tit As Titulo, ByVal CodEmpresa As String, ByVal EndEmpresa As String, ByVal Conta As String, ByVal CodCliente As String, ByVal EndCliente As String, ByVal TemCliente As Boolean, ByVal DEBCRE As String, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, Optional ByVal Conciliado As Boolean = False, Optional ByVal DataConciliacao As String = "")
        Dim Historico As String = String.Empty

        sql = "INSERT INTO Razao " & vbCrLf & _
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
              "       Conciliacao, " & vbCrLf & _
              "       DataDaBaixa, " & vbCrLf & _
              "       Historico, " & vbCrLf & _
              "       PrevistoRealizado," & vbCrLf & _
              "       Processo," & vbCrLf & _
              "       UsuarioInclusao," & vbCrLf & _
              "       UsuarioInclusaoData)" & vbCrLf & _
              "VALUES ('" & CodEmpresa & "'," & vbCrLf & _
              "         " & EndEmpresa & "," & vbCrLf & _
              "        '" & Conta & "'" & vbCrLf

        If TemCliente Then
            sql &= ", '" & CodCliente & "'"  'Cliente
            sql &= ", " & EndCliente         'Endereco do Cliente
        Else
            sql &= ", ''"                    'Cliente
            sql &= ", 0"                     'Endereco do Cliente
        End If

        sql &= ", '" & tit.Baixa.ToString("yyyy/MM/dd") & "'" 'Data de Movimento
        sql &= ", 0070"
        sql &= ", " & tit.Codigo                          'Sequencia no Razao = Registro do Titulo
        sql &= ", " & tit.Codigo                          'Numero do Titulo
        sql &= ", '" & tit.CodigoUnidadeDeNegocio & "'"   'Unidade de Negócio
        sql &= ", " & tit.CodigoIndexador                 'Indexador
        sql &= ", '" & tit.DataMoeda.ToString("yyyy/MM/dd") & "'" 'Data da Moeda

        'Valor Oficial
        If DEBCRE = "D" Then
            sql &= ", " & Str(ValorOficial)  'Valor Débito Oficial
            sql &= ", 0.0"                   'Valor Crédito Oficial
        Else
            sql &= ", 0.0"                   'Valor Debito Oficial
            sql &= ", " & Str(ValorOficial)  'Valor Crédito Oficial
        End If

        If IsVariacao(CodEmpresa, EndEmpresa, Conta) Then
            sql &= ", 0.0"      'Valor Débito Moeda
            sql &= ", 0.0"      'Valor Crédito Moeda
        Else
            If DEBCRE = "D" Then
                sql &= ", " & Str(ValorMoeda) 'Valor Débito Moeda
                sql &= ", 0.0"                'Valor Crédito Moeda
            Else
                sql &= ", 0.0"                'Valor Debito Moeda
                sql &= ", " & Str(ValorMoeda) 'Valor Crédito Moeda
            End If
        End If

        If DEBCRE = "C" AndAlso Conciliado Then
            sql &= ", 'B'"                                              'Conciliação
            sql &= ", '" & CDate(DataConciliacao).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
        Else
            sql &= ", NULL "                                            'Conciliação
            sql &= ", NULL "                                            'Data Conciliação
        End If

        If tit.NumeroDoCheque > 0 AndAlso Left(Conta, 7) = "1010102" Then
            Historico = "CH.NR. " & tit.NumeroDoCheque & " - " & tit.Cliente.Nome
        Else
            Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)
        End If

        sql &= ",'" & Historico & "'" 'Histórico

        sql &= ", 'P'"                         'Previsto/Realizado
        If tit.ReceberPagar = "P" Then
            sql &= ", 'CONTASAPAGAR'"          'Processo
        Else
            sql &= ", 'CONTASARECEBER'"        'Processo
        End If
        sql &= ", '" & tit.UsuarioBaixa & "'"                             'Usuario que Baixou
        sql &= ", '" & tit.UsuarioBaixaData.ToString("yyyy/MM/dd") & "')" 'Data da Baixa

        Sqls.Add(sql)

        strHTMLraz &= "<tr>" & vbCrLf & _
                   "<td><A Class='A5'>" & CodEmpresa & "-" & EndEmpresa & "</A></td>" & vbCrLf & _
                   "<td><A Class='A5'>" & Conta & "</A></td>" & vbCrLf

        If TemCliente Then
            strHTMLraz &= "<td><A Class='A5'>" & CodCliente & "-" & EndCliente & "</A></td>" & vbCrLf
        Else
            strHTMLraz &= "<td><A Class='A5'></A></td>" & vbCrLf
        End If

        strHTMLraz &= "<td><A Class='A5'>" & tit.Baixa.ToString("yyyy/MM/dd") & "</A></td>" & vbCrLf & _
                   "<td><A Class='A5'>" & tit.Codigo & "</A></td>" & vbCrLf

        If DEBCRE = "D" Then
            totalDebito += ValorOficial
            strHTMLraz &= "<td align='right'><A Class='A5'>" & Str(ValorOficial) & "</A></td>" & vbCrLf & _
                       "<td align='right'><A Class='A5'>0,00</A></td>" & vbCrLf
        Else
            totalCredito += ValorOficial
            strHTMLraz &= "<td align='right'><A Class='A5'>0,00</A></td>" & vbCrLf & _
                   "<td align='right'><A Class='A5'>" & Str(ValorOficial) & "</A></td>" & vbCrLf
        End If

        strHTMLraz &= "<td><A Class='A5'>" & Historico & "</A></td>" & vbCrLf & _
                      "<td><A Class='A5'>" & tit.UsuarioBaixa & "</A></td>" & vbCrLf & _
                      "<td><A Class='A5'>" & tit.UsuarioBaixaData.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                      "</tr>" & vbCrLf
    End Sub

    Private Function IsVariacao(ByVal Empresa As String, EndEmpresa As String, ByVal Conta As String) As Boolean
        Dim sql As String = "SELECT CASE" & vbCrLf & _
                            "        WHEN '" & Conta & "' in (empDed.contaVariacaoMonetariaAtiva, empDed.contaVariacaoMonetariaPassiva) " & vbCrLf & _
                            "          THEN 1" & vbCrLf & _
                            "          ELSE 0" & vbCrLf & _
                            "       END Variacao" & vbCrLf & _
                            "  FROM clientesxempresas empDed" & vbCrLf & _
                            " Where empDed.empresa_id    = '" & Empresa & "'" & vbCrLf & _
                            "   AND empDed.endempresa_id =  " & EndEmpresa

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

    Protected Sub lnkVincular_Click(sender As Object, e As EventArgs) Handles lnkVincular.Click
        If gridParcelas.Rows.Count = 0 Then
            MsgBox(Me.Page, "Parcelamento do Pedido não foi feito!")
            Exit Sub
        End If

        Dim processoFinalizado As Boolean = False

        Dim NomeArquivo As String
        Dim NomeArquivo2 As String = "Files/resumopedido" & txtPedido.Text & ".html"
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/resumopedido" & txtPedido.Text & ".html"))

        Dim strHTML As String = "<html>" & vbCrLf & _
                            "   <head>" & vbCrLf & _
                            "     <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
                            "       <title>RESUMO DA CONVERSÃO DO PEDIDO</title>" & vbCrLf & _
                            "       <style type=""text/css"">" & vbCrLf & _
                            "           H6 {page-break-after:always}" & vbCrLf & _
                            "           A.A1 {font-family: Ms serif;font-size: 5pt;line height: 15px}" & vbCrLf & _
                            "           A.A2 {font-family: Ms serif;font-size: 6pt;line height: 15px}" & vbCrLf & _
                            "           A.A3 {font-family: Ms serif;font-size: 7pt;line height: 15px}" & vbCrLf & _
                            "           A.A4 {font-family: Ms serif;font-size: 8pt;line height: 15px}" & vbCrLf & _
                            "           A.A5 {font-family: Ms serif;font-size: 9pt;line height: 15px}" & vbCrLf & _
                            "           A.A6 {font-family: Ms serif;font-size: 10pt;line height: 15px}" & vbCrLf & _
                            "           A.A7 {font-family: Ms serif bold;font-size: 8pt;line height: 15px}" & vbCrLf & _
                            "           A.A8 {font-family: Ms serif bold;font-size: 9pt;line height: 15px}" & vbCrLf & _
                            "           A.A9 {font-family: Ms serif bold;font-size: 10pt;line height: 15px}" & vbCrLf & _
                            "           A.A10 {font-family: Ms serif bold;font-size: 11pt;line height: 15px}" & vbCrLf & _
                            "           A.A11 {font-family: Ms serif bold;font-size: 12pt;line height: 15px}" & vbCrLf & _
                            "           A.A12 {font-family: Ms serif bold;font-size: 13pt;line height: 15px}" & vbCrLf & _
                            "           A.A13 {font-family: Ms serif bold;font-size: 14pt;line height: 15px}" & vbCrLf & _
                            "       </style>" & vbCrLf & _
                            "   </head>" & vbCrLf & _
                            "   <bodyb text=#000000 bgcolor=#FFFFFF>"

        strHTML &= "<table border='0' width='100%'>" & vbCrLf & _
                   "<tr>" & vbCrLf & _
                   "<td>" & vbCrLf

        Try
            RecuperarPedido()

            Dim Sqls As New ArrayList

            'Cria parcelas no PedidoXParcela
            strHTML &= "<table border='0'>" & vbCrLf & _
                        "<tr>" & vbCrLf & _
                        "<td colspan='3'>" & vbCrLf & _
                        "<hr />" & vbCrLf & _
                        "</td>" & vbCrLf & _
                        "</tr>" & vbCrLf & _
                        "<tr>" & vbCrLf & _
                        "<td colspan='3'><A Class='A10'>PEDIDOxPARCELA</A></td>" & vbCrLf & _
                        "</tr>" & vbCrLf & _
                        "<tr>" & vbCrLf & _
                        "<td><A Class='A5'>PARCELA</A></td>" & vbCrLf & _
                        "<td><A Class='A5'>VENCIMENTO</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                        "</tr>" & vbCrLf

            For Each pRow As GridViewRow In gridParcelas.Rows
                Dim parcela As New PedidoXParcela(objPedido)
                parcela.IUD = "I"
                parcela.CodigoParcela = pRow.Cells(1).Text
                parcela.DataVencimento = pRow.Cells(2).Text
                parcela.Valor = CDec(pRow.Cells(3).Text)
                parcela.SalvarSql(Sqls)

                strHTML &= "<tr>" & vbCrLf & _
                          "<td><A Class='A5'>" & parcela.CodigoParcela & "</A></td>" & vbCrLf & _
                          "<td><A Class='A5'>" & parcela.DataVencimento & "</A></td>" & vbCrLf & _
                          "<td align='right'><A Class='A5'>" & parcela.Valor & "</A></td>" & vbCrLf & _
                          "</tr>" & vbCrLf
            Next

            strHTML &= "<tr>" & vbCrLf & _
                      "<td colspan='3'>" & vbCrLf & _
                      "<hr />" & vbCrLf & _
                      "</td>" & vbCrLf & _
                      "</tr>" & vbCrLf & _
                      "</table>" & vbCrLf

            strHTML &= "<table border='0'>" & vbCrLf & _
                        "<tr>" & vbCrLf & _
                        "<td colspan='12'>" & vbCrLf & _
                        "<hr />" & vbCrLf & _
                        "</td>" & vbCrLf & _
                        "</tr>" & vbCrLf & _
                        "<tr>" & vbCrLf & _
                        "<td colspan='12'><A Class='A10'>VENCIMENTOS NOVOS</A></td>" & vbCrLf & _
                        "</tr>" & vbCrLf & _
                        "<tr>" & vbCrLf & _
                        "<td align='center'><A Class='A5'>PAG/REC</A></td>" & vbCrLf & _
                        "<td><A Class='A5'>TITULO</A></td>" & vbCrLf & _
                        "<td align='center'><A Class='A5'>PROVISÃO</A></td>" & vbCrLf & _
                        "<td><A Class='A5'>VENCIMENTO</A></td>" & vbCrLf & _
                        "<td><A Class='A5'>BAIXA</A></td>" & vbCrLf & _
                        "<td align='center'><A Class='A5'>MOEDA</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>VALOR</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>DEDUÇÃO</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>DESCONTO</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>ACRÉSCIMOS</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>JUROS</A></td>" & vbCrLf & _
                        "<td align='right'><A Class='A5'>LÍQUIDO</A></td>" & vbCrLf & _
                        "</tr>" & vbCrLf

            For Each nRow As GridViewRow In GridNotasXTitulos.Rows
                Dim titulo As New Titulo(nRow.Cells(5).Text, nRow.Cells(3).Text)

                Dim nf As New [Lib].Negocio.NotaFiscal()

                For Each rowNF As GridViewRow In GridNotas.Rows
                    If nRow.Cells(0).Text = rowNF.Cells(6).Text AndAlso nRow.Cells(1).Text = rowNF.Cells(7).Text AndAlso nRow.Cells(2).Text = rowNF.Cells(8).Text Then
                        nf.CodigoEmpresa = rowNF.Cells(1).Text
                        nf.EnderecoEmpresa = rowNF.Cells(2).Text
                        nf.CodigoCliente = rowNF.Cells(3).Text
                        nf.EnderecoCliente = rowNF.Cells(4).Text
                        nf.EntradaSaida = IIf(rowNF.Cells(6).Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                        nf.Codigo = rowNF.Cells(8).Text
                        nf.Serie = rowNF.Cells(7).Text
                        nf = New [Lib].Negocio.NotaFiscal(nf)
                        Exit For
                    End If
                Next

                If titulo.Codigo = 0 Then
                    Dim Tit As New Titulo

                    'NUMERADOR DOS TITULOS
                    Dim SqlN As String = "exec sp_Numerador '" & HttpContext.Current.Session("ssNomeServidor") & "',0," & [Lib].Negocio.eTiposNumerador.Titulo & ""
                    Dim dsN As New DataSet
                    dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                    Dim CodigoNumerador As Integer = 0
                    If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                        CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                    End If

                    If Not CodigoNumerador > 0 Then
                        MsgBox(Me.Page, "Numerador de Títulos não cadastrado!")
                        Exit Sub
                    End If

                    Tit.IndiceFixo = False

                    Tit.IUD = "I"
                    Tit.Codigo = CodigoNumerador
                    Tit.Sequencia = 0

                    'Caso Pedido AFixar vire PROVISAO = 4 (AFIXAR) - FURLAN - 04/10/2016
                    If objPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
                        Tit.CodigoProvisao = 4
                    Else
                        Tit.CodigoProvisao = 2
                    End If

                    If nf.EntradaSaida = eEntradaSaida.Entrada Then
                        Tit.ReceberPagar = "P"
                        If nf.SubOperacao.Devolucao Then
                            Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                            Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraVenda
                        Else
                            Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                            Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraCompra
                        End If
                    Else
                        Tit.ReceberPagar = "R"
                        If nf.SubOperacao.Devolucao Then
                            Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                            Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraCompra
                        Else
                            Tit.ContaContabilCliente = nf.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                            Tit.CodigoCarteira = nf.Itens(0).Produto.CodigoCarteiraVenda
                        End If
                    End If

                    Tit.Tributo = ""

                    Tit.CodigoEmpresaPedido = nf.Pedido.CodigoEmpresa
                    Tit.EndEmpresaPedido = nf.Pedido.EnderecoEmpresa
                    Tit.CodigoPedido = nf.CodigoPedido
                    Tit.CodigoPedidoFixacao = 0
                    Tit.CodigoProcuracao = nf.CodigoProcuracao
                    Tit.DataMoeda = CDate(nRow.Cells(6).Text) ' nf.Movimento.ToString("dd-MM-yyyy")
                    Tit.CodigoIndexador = nf.Pedido.CodigoIndexador
                    Tit.CodigoMoeda = nf.Pedido.CodigoMoeda
                    Tit.Movimento = nf.Movimento
                    Tit.Vencimento = Funcoes.ValidaDataUtil(nf.Empresa.Codigo, nf.Empresa.CodigoEndereco, CDate(nRow.Cells(6).Text))

                    Tit.Prorrogacao = Funcoes.ValidaDataUtil(nf.Empresa.Codigo, nf.Empresa.CodigoEndereco, CDate(nRow.Cells(6).Text))
                    Tit.Baixa = Funcoes.ValidaDataUtil(nf.Empresa.Codigo, nf.Empresa.CodigoEndereco, CDate(nRow.Cells(6).Text))
                    Tit.CodigoTipoPgto = 1
                    Tit.CodigoSituacao = 1
                    Tit.Lote = 70
                    Tit.CodigoUnidadeDeNegocio = nf.Pedido.CodigoUnidadeNegocio
                    Tit.CodigoEmpresa = nf.CodigoEmpresa
                    Tit.EnderecoEmpresa = nf.EnderecoEmpresa
                    Tit.CodigoCliente = nf.Pedido.CodigoCliente
                    Tit.EndCliente = nf.Pedido.EnderecoCliente
                    Tit.CodigoBancoCliente = 0
                    Tit.CodigoAgenciaCliente = ""
                    Tit.DigitoAgenciaCliente = ""
                    Tit.ContaCliente = ""
                    Tit.DigitoContaCliente = ""
                    Tit.Cheque = False
                    Tit.Slips = False
                    Tit.Recibo = False
                    Tit.Aviso = False
                    Tit.ReciboDeposito = False

                    Tit.ValorDoDocumento = CDec(nRow.Cells(7).Text)

                    If nf.Codigo > 0 AndAlso nf.Pedido.CodigoMoeda = 3 Then
                        Tit.MoedaValorDoDocumento = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                        If Tit.MoedaValorDoDocumento = 0 Then
                            MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                            Exit Sub
                        End If
                    End If

                    Tit.Historico = "REF. NF " & nf.Codigo & "-" & nf.Serie & "-" & IIf(nf.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Pedido: " & nf.CodigoPedido & " / " & nf.Cliente.Nome
                    Tit.CodigoDeBarras = ""
                    Tit.CodigoDigitado = False
                    Tit.CodigoDeBarrasPreImpresso = False

                    Tit.CodigoDestinatario = nf.Pedido.CodigoCliente
                    Tit.EndDestinatario = nf.Pedido.EnderecoCliente
                    Tit.NomeDoDestinatario = ""
                    Tit.Destinacao = ""
                    Tit.Solicitacao = 0

                    Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
                    Tit.RegistroMestre = 0
                    Tit.Observacoes = ""
                    Tit.SituacaoBancaria = 0
                    Tit.NumeroDoCheque = 0
                    Tit.CodigoAdiantamento = 0
                    Tit.TaxaAdto = 0

                    Tit.UsuarioInclusao = nf.UsuarioInclusao
                    Tit.UsuarioInclusaoData = nf.DataInclusao

                    Tit.SalvarSql(Sqls)

                    strHTML &= "<tr>" & vbCrLf & _
                               "<td align='center'><A Class='A5'>" & Tit.ReceberPagar & "</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>" & Tit.Codigo & "</A></td>" & vbCrLf & _
                               "<td align='center'><A Class='A5'>" & Tit.CodigoProvisao & "</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>" & Tit.Prorrogacao.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                               "<td><A Class='A5'>" & Tit.Baixa.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf
                    If Tit.CodigoMoeda = 1 Then
                        strHTML &= "<td align='center'><A Class='A5'>R$</A></td>" & vbCrLf
                    Else
                        strHTML &= "<td align='center'><A Class='A5'>U$</A></td>" & vbCrLf
                    End If

                    strHTML &= "<td align='right'><A Class='A5'>" & Tit.ValorDoDocumento & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & Tit.Deducoes & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & Tit.Descontos & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & Tit.Acrescimos & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & Tit.Juros & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & Tit.ValorLiquido & "</A></td>" & vbCrLf & _
                              "</tr>" & vbCrLf

                    'RELACIONAR NOTA FISCAL COM NOVO TÍTULO
                    sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf & _
                          " values('" & Tit.CodigoEmpresaPedido & "'," & Tit.EndEmpresaPedido & ", " & vbCrLf & _
                          "'" & Tit.CodigoCliente & "'," & Tit.EndCliente & "," & vbCrLf & _
                          "'" & nRow.Cells(0).Text & "','" & nRow.Cells(1).Text & "'," & nRow.Cells(2).Text & "," & Tit.Codigo & ")"

                    Sqls.Add(sql)
                ElseIf titulo.CodigoProvisao = 2 Then
                    'NUMERADOR DOS TITULOS
                    Dim SqlN As String = "exec sp_Numerador '" & HttpContext.Current.Session("ssNomeServidor") & "',0," & [Lib].Negocio.eTiposNumerador.Titulo & ""
                    Dim dsN As New DataSet
                    dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                    Dim CodigoNumerador As Integer = 0
                    If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                        CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                    End If

                    If Not CodigoNumerador > 0 Then
                        MsgBox(Me.Page, "Numerador de Títulos não cadastrado!")
                        Exit Sub
                    End If

                    'GRAVAR NOVO TÍTULO
                    titulo.IUD = "I"
                    titulo.Codigo = CodigoNumerador
                    titulo.ValorDoDocumento = CDec(nRow.Cells(7).Text)

                    If nf.Codigo > 0 AndAlso nf.Pedido.CodigoMoeda = 3 Then
                        titulo.MoedaValorDoDocumento = nf.Itens.Sum(Function(s) s.ValorLiquidoMoeda)

                        If titulo.MoedaValorDoDocumento = 0 Then
                            MsgBox(Me.Page, "Valor em Dólar da Nota Fiscal " & nf.Codigo & " não foi encontrado. Entre em contato com o suporte.")
                            Exit Sub
                        End If
                    End If

                    titulo.SalvarSql(Sqls, False)

                    strHTML &= "<tr>" & vbCrLf & _
                              "<td align='center'><A Class='A5'>" & titulo.ReceberPagar & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & titulo.Codigo & "</A></td>" & vbCrLf & _
                              "<td align='center'><A Class='A5'>" & titulo.CodigoProvisao & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & titulo.Prorrogacao.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf & _
                              "<td><A Class='A5'>" & titulo.Baixa.ToString("dd/MM/yyyy") & "</A></td>" & vbCrLf
                    If titulo.CodigoMoeda = 1 Then
                        strHTML &= "<td align='center'><A Class='A5'>R$</A></td>" & vbCrLf
                    Else
                        strHTML &= "<td align='center'><A Class='A5'>U$</A></td>" & vbCrLf
                    End If

                    strHTML &= "<td align='right'><A Class='A5'>" & titulo.ValorDoDocumento & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & titulo.Deducoes & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & titulo.Descontos & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & titulo.Acrescimos & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & titulo.Juros & "</A></td>" & vbCrLf & _
                              "<td align='right'><A Class='A5'>" & titulo.ValorLiquido & "</A></td>" & vbCrLf & _
                              "</tr>" & vbCrLf

                    'RELACIONAR NOTA FISCAL COM NOVO TÍTULO
                    sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf & _
                          " values('" & titulo.CodigoEmpresaPedido & "'," & titulo.EndEmpresaPedido & ", " & vbCrLf & _
                          "'" & titulo.CodigoCliente & "'," & titulo.EndCliente & "," & vbCrLf & _
                          "'" & nRow.Cells(0).Text & "','" & nRow.Cells(1).Text & "'," & nRow.Cells(2).Text & "," & titulo.Codigo & ")"

                    Sqls.Add(sql)
                End If
            Next

            'Cancela apenas os Titulos em aberto 
            For Each row As GridViewRow In gridFinanceiro.Rows
                If row.Cells(2).Text = "2" Or row.Cells(2).Text = "3" Then
                    If row.Cells(0).Text = "P" Then
                        sql = "Update ContasAPagar set Situacao = 3 where Registro_id = " & row.Cells(1).Text
                    Else
                        sql = "Update ContasAReceber set Situacao = 3 where Registro_id = " & row.Cells(1).Text
                    End If
                    Sqls.Add(sql)
                End If
            Next

            'Atualiza Momento Financeiro do Pedido 
            sql = "Update Pedidos set MomentoFinanceiro = 9, CondicaoPagamento = " & ddlCondPagPed.SelectedValue & vbCrLf & _
                "where Empresa_id = '" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                "  and EndEmpresa_id = " & objPedido.EnderecoEmpresa & vbCrLf & _
                "  and Pedido_id = " & objPedido.Codigo
            Sqls.Add(sql)

            If Banco.GravaBanco(Sqls) Then
                processoFinalizado = True
                HiddenArquivo.Value = "resumopedido" & txtPedido.Text & ".html"
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            strHTML &= "</td>" & vbCrLf & _
                      "</tr>" & vbCrLf & _
                      "</table>" & vbCrLf & _
                      "</body>" & vbCrLf & _
                      "</html>"

            strm.WriteLine(strHTML)

            strm.Close()
            strm.Dispose()

            If processoFinalizado Then
                Limpar(False)
                lnkConsultar_Click(Nothing, Nothing)

                cmdArquivoDeSaida.Visible = True

                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
            End If
        End Try
    End Sub

    Protected Sub ddlCondPagPed_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCondPagPed.SelectedIndexChanged
        Try
            RecuperarPedido()
            objPedido.CodigoCondicaoPagamento = ddlCondPagPed.SelectedValue
            SalvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkParcelar_Click(sender As Object, e As EventArgs) Handles LnkParcelar.Click
        RecuperarPedido()
        If Not IsDate(txtDataCondPagParcela.Text) Then
            MsgBox(Me.Page, "Informe a Data inicial do Parcelamento")
            Exit Sub
        ElseIf CDate(txtDataCondPagParcela.Text).ToString("yyyy/MM/dd") < objPedido.DataPedido.ToString("yyyy/MM/dd") Then
            MsgBox(Me.Page, "Data do inicio do parcelamento nao pode ser menor que a data do Pedido")
            Exit Sub
        ElseIf objPedido.CodigoCondicaoPagamento = 0 Then
            MsgBox(Me.Page, "Condição de Pagamento não foi selecionada")
            Exit Sub
        End If

        objPedido.Parcelas.Parcelar(CDate(txtDataCondPagParcela.Text))
        gridParcelas.DataSource = objPedido.Parcelas
        gridParcelas.DataBind()

        txtPedidoTotal.Text = objPedido.Parcelas.Sum(Function(s) s.Valor).ToString("N2")

        Dim totalbaixado As Decimal = 0
        For Each col As GridViewRow In GridNotasXTitulos.Rows
            If col.Cells(5).Text = "1" Then
                If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If col.Cells(0).Text = "E" Then
                        totalbaixado += col.Cells(8).Text
                    Else
                        totalbaixado -= col.Cells(8).Text
                    End If
                Else
                    If col.Cells(0).Text = "S" Then
                        totalbaixado += col.Cells(8).Text
                    Else
                        totalbaixado -= col.Cells(8).Text
                    End If
                End If
            End If
        Next

        txtPedidoTotalPago.Text = totalbaixado.ToString("N2")

        txtPedidoSaldo.Text = (CDec(txtPedidoTotal.Text) - CDec(txtPedidoTotalPago.Text)).ToString("N2")
    End Sub

    Protected Sub gridParcelas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridParcelas.SelectedIndexChanged
        Try
            RecuperarPedido()

            txtCodigoParcela.Text = objPedido.Parcelas(gridParcelas.SelectedIndex).CodigoParcela
            txtDataVencParcela.Text = objPedido.Parcelas(gridParcelas.SelectedIndex).DataVencimento.ToString("dd/MM/yyyy")
            txtValorParcela.Text = objPedido.Parcelas(gridParcelas.SelectedIndex).Valor.ToString("N2")

            If gridParcelas.Rows.Count = 1 Or gridParcelas.SelectedIndex + 1 = gridParcelas.Rows.Count Then
                txtValorParcela.Enabled = False
            Else
                txtValorParcela.Enabled = True
            End If

            If Not String.IsNullOrWhiteSpace(objPedido.Contrato) Then
                txtValorParcela.Enabled = False
                MsgBox(Me.Page, "Valor da Parcela com Contrato Bancário não pode ser alterado, qualquer dúvida entre em contato com o Financeiro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkAtualizarParcela_Click(sender As Object, e As EventArgs) Handles LnkAtualizarParcela.Click
        If gridParcelas.Rows.Count = 0 Then Exit Sub
        Dim nPar As Integer = CInt(txtCodigoParcela.Text)
        If Not IsNumeric(txtValorParcela.Text) Then
            MsgBox(Me.Page, "Informe um valor para a Parcela.")
            Exit Sub
        End If

        RecuperarPedido()
        If objPedido.Parcelas.Count > nPar Then
            If objPedido.Parcelas(nPar - 1).DataVencimento > objPedido.Parcelas(nPar).DataVencimento Then
                MsgBox(Me.Page, "A data da Parcela atual nao pode ser maior que a data da parcela subsequente.")
                Exit Sub
            End If
        End If

        objPedido.Parcelas.ModificarParcela(nPar, CDate(txtDataVencParcela.Text), CDec(txtValorParcela.Text))

        If objPedido.Parcelas.MSG.Length > 0 Then
            MsgBox(Me.Page, objPedido.Parcelas.MSG)
            objPedido.Parcelas.MSG = ""
            Exit Sub
        End If

        SalvarPedido()
        gridParcelas.DataSource = objPedido.Parcelas
        gridParcelas.DataBind()

        txtCodigoParcela.Text = String.Empty
        txtDataVencParcela.Text = String.Empty
        txtValorParcela.Text = String.Empty
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            ConsultaRegistros()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LiberaPedidoVirtual")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class