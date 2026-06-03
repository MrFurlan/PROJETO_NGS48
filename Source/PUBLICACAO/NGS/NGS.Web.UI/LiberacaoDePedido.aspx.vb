Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LiberacaoDePedido
    Inherits BasePage

    Private objPedido As [Lib].Negocio.Pedido
    Private sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("LiberacaoDePedido", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                CargaCondicoesPagamento()
                CarregaUsuarios()

                Limpar(True)
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

    Private Sub CarregaUsuarios()
        ddl.Carregar(ddlUsuarioLiberacao, CarregarDDL.Tabela.Usuarios, "", True)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub CargaCondicoesPagamento()
        Dim objCondicoes As New [Lib].Negocio.ListCondicaoPagamento(True)

        For Each objCondicao As [Lib].Negocio.CondicaoPagamento In objCondicoes
            lstCondicoes.Items.Add(New ListItem(objCondicao.Codigo.ToString() & " " & _
                                                Funcoes.AlinharEsquerda(objCondicao.Descricao.Trim(), 50, ".") & " " & _
                                                Funcoes.AlinharDireita(objCondicao.Parcelas.ToString(), 2, " "), _
                                                objCondicao.Codigo.ToString()))

            lstCondicoesPgtoEntrega.Items.Add(New ListItem(objCondicao.Codigo.ToString() & " " & _
                                                Funcoes.AlinharEsquerda(objCondicao.Descricao.Trim(), 50, ".") & " " & _
                                                Funcoes.AlinharDireita(objCondicao.Parcelas.ToString(), 2, " "), _
                                                objCondicao.Codigo.ToString()))
        Next
        Funcoes.InserirLinhaEmBranco(lstCondicoes)
        Funcoes.InserirLinhaEmBranco(lstCondicoesPgtoEntrega)
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

    Private Sub Limpar(ByVal Geral As Boolean)

        If Geral Then
            txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
            txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            gridConsulta.DataSource = Nothing
            gridConsulta.DataBind()
            TabContainer1.ActiveTabIndex = 0
        End If

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
        txtIndiceFixado.Text = String.Empty
        txtProduto.Text = String.Empty
        txtQuantidade.Text = String.Empty
        txtUnitario.Text = String.Empty
        txtValor.Text = String.Empty
        txtUnitarioMoeda.Text = String.Empty
        txtValorMoeda.Text = String.Empty
        txtNaturezaDaOperacao.Text = String.Empty

        gridEncargos.DataSource = Nothing
        gridEncargos.DataBind()
        gridEncargos.Parent.Visible = False

        gridTransportes.DataSource = Nothing
        gridTransportes.DataBind()
        gridTransportes.Parent.Visible = False

        gridRepresentantes.DataSource = Nothing
        gridRepresentantes.DataBind()
        gridRepresentantes.Parent.Visible = False

        txtCifFob.Text = String.Empty

        ddlMomentoFinanceiro.Enabled = False
        txtDataVencimento.Text = String.Empty
        txtValorVencimento.Text = String.Empty
        lstCondicoes.SelectedIndex = 0
        lstCondicoes.Enabled = False
        btnOkVencimento.Enabled = False
        txtValorVencimento.Enabled = False
        txtDataVencimento.Enabled = False
        gridFinanceiro.DataSource = Nothing
        gridFinanceiro.DataBind()
        ajustarFinanceiro.Value = False
        lstCondicoesPgtoEntrega.SelectedIndex = 0
        txtQuotaDeEntrega.Text = String.Empty
        ddlPeriodicidadeEntrega.SelectedIndex = 0
        btnAjustarEntrega.Enabled = True
        pnlEntrega.Visible = False
        Session.Remove("RowIndex" & HID.Value)
        Session.Remove("objClienteLibPed" & HID.Value)
        Session.Remove("objPedido" & HID.Value)
        Session.Remove("objPedVencimentos" & HID.Value)
        HID.Value = (New Random).Next
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
        If Funcoes.VerificaPermissao("LiberacaoDePedido", "LEITURA") Then
            Dim Quimica As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "05272759"
            Dim Agricola As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "04854422"
            Dim Fex As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "15204808"
            Dim Nutri As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "05366261"
            Dim NutriLog As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "38198213"
            Dim Baxi As Boolean = (Left(UsuarioServidor.CodigoEmpresa, 8) = "40938762" OrElse Left(UsuarioServidor.CodigoEmpresa, 8) = "49673784")
            Dim Verde As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "44979506"
            Dim RTGraos As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "24450490"

            sql = "select p.Empresa_id, p.EndEmpresa_Id, cliEmp.Nome, cliEmp.Cidade, p.Pedido_Id, P.PedidoEfetivo, p.Cliente, p.EndCliente, cli.Nome as NomeCliente, " & vbCrLf & _
                  "		p.DataPedido, p.Moeda, cope.Classe_id, " & IIf(Quimica, "prd.Grupo as Produto_Id, gp.Descricao as NomeProduto," _
                                                                       , "pXi.Produto_Id, prd.Nome as NomeProduto,") & "isnull(p.PedidoBloqueado,0) AS PedidoBloqueado, " & vbCrLf & _
                  IIf(Quimica, "Sum(pXi.Quantidade) AS Quantidade,", "pXi.Quantidade,") & vbCrLf & _
                  IIf(Quimica, "Sum(", "") & "	case " & vbCrLf & _
                  "			                      when p.Moeda = 1 " & vbCrLf & _
                  "				                   then pXi.UnitarioOficial " & vbCrLf & _
                  "				                   else pXi.UnitarioMoeda " & vbCrLf & _
                  "		                         end" & IIf(Quimica, ")", "") & " as Unitario, " & vbCrLf & _
                  IIf(Quimica, "Sum(", "") & "		case " & vbCrLf & _
                  "			                          when p.Moeda = 1 " & vbCrLf & _
                  "				                       then pXi.TotalOficial " & vbCrLf & _
                  "				                       else pXi.TotalMoeda " & vbCrLf & _
                  "		                            end" & IIf(Quimica, ")", "") & " as Total, " & vbCrLf & _
                  "		                            P.UsuarioLiberacao," & vbCrLf & _
                  "					CASE " & vbCrLf & _
                  "					   WHEN P.UsuarioLiberacao <> '' " & vbCrLf & _
                  "					     THEN P.UsuarioLiberacaoData" & vbCrLf & _
                  "                      Else NULL" & vbCrLf & _
                  "					END as UsuarioLiberacaoData " & vbCrLf & _
                  " from Pedidos p " & vbCrLf & _
                  "     inner join (Select pedI.Empresa_id, pedI.EndEmpresa_Id, pedI.Pedido_Id, pedI.Produto_Id, " & vbCrLf & _
                  "                     sum(Case  " & vbCrLf & _
                  "                           when pedI.TipoDeLancamento = 'N'  " & vbCrLf & _
                  "                            then pedI.UnitarioOficial  " & vbCrLf & _
                  "                            else 0  " & vbCrLf & _
                  "                         end) UnitarioOficial,  " & vbCrLf & _
                  "                     sum(Case  " & vbCrLf & _
                  "                           when pedI.TipoDeLancamento = 'N'  " & vbCrLf & _
                  "                            then pedI.UnitarioMoeda  " & vbCrLf & _
                  "                            else 0  " & vbCrLf & _
                  "                         end) UnitarioMoeda,  " & vbCrLf & _
                  "                     sum(case   " & vbCrLf & _
                  " 						  when pedI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  " 							then pedI.Quantidade * -1  " & vbCrLf & _
                  " 							else pedI.Quantidade  " & vbCrLf & _
                  " 						end) as Quantidade,  " & vbCrLf & _
                  " 				    sum(case   " & vbCrLf & _
                  " 						  when pedI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  " 							then pedI.TotalOficial * -1   " & vbCrLf & _
                  " 							else pedI.TotalOficial  " & vbCrLf & _
                  " 						end) as TotalOficial,  " & vbCrLf & _
                  " 				    sum(case   " & vbCrLf & _
                  " 						  when pedI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  " 							then pedI.TotalMoeda * -1   " & vbCrLf & _
                  " 							else pedI.TotalMoeda  " & vbCrLf & _
                  " 						end) as TotalMoeda  " & vbCrLf & _
                  "				from PedidoXItemxLancamento pedI " & vbCrLf & _
                  "					inner join Pedidos ped " & vbCrLf & _
                  "							on ped.Empresa_Id     = pedI.Empresa_Id " & vbCrLf & _
                  "							and ped.EndEmpresa_Id = pedI.EndEmpresa_Id " & vbCrLf & _
                  "							and ped.Pedido_Id     = pedI.Pedido_Id " & vbCrLf & _
                  "				where ped.Situacao = 1 " & vbCrLf & _
                  "				group by pedI.Empresa_Id, pedI.EndEmpresa_Id, pedI.Pedido_Id, pedI.Produto_Id) pXi " & vbCrLf & _
                  "			on pXi.Empresa_Id = p.Empresa_Id " & vbCrLf & _
                  "			and pXi.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf & _
                  "			and pXi.Pedido_Id = p.Pedido_Id " & vbCrLf & _
                  "	    inner join Produtos prd " & vbCrLf & _
                  "		    	 on prd.Produto_id = pXi.Produto_id " & vbCrLf

            If Not Quimica And Not Agricola And Not Fex And Not Nutri And Not NutriLog And Not Baxi And Not Verde And Not RTGraos Then
                sql &= "  and prd.Agrupar    = 'N' " & vbCrLf
            Else
                sql &= "inner Join GruposDeEstoques gp" & vbCrLf & _
                       "    on gp.Grupo_Id = prd.grupo" & vbCrLf
            End If

            sql &= "	    inner join SubOperacoes so " & vbCrLf & _
                  "		    	 on so.Operacao_id     = p.Operacao " & vbCrLf & _
                  "			    and so.SubOperacoes_id = p.Suboperacao " & vbCrLf & _
                  "	    inner join ClassesDeOperacoes cope " & vbCrLf & _
                  "		    	on cope.Classe_id = so.Classe " & vbCrLf & _
                  "	    inner join Clientes cliEmp " & vbCrLf & _
                  "		    	 on cliEmp.Cliente_id  = p.Empresa_Id " & vbCrLf & _
                  "		    	and cliEmp.Endereco_id = p.EndEmpresa_Id " & vbCrLf & _
                  "	    inner join Clientes cli " & vbCrLf & _
                  "		    	 on cli.Cliente_id  = p.Cliente " & vbCrLf & _
                  "		    	and cli.Endereco_id = p.EndCliente " & vbCrLf & _
                  " where p.Situacao = 1 " & vbCrLf & _
                  "  and cope.Classe_id in('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.COMPRASAORDEM.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.LUCROREAL.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "', '" & eClassesOperacoes.AFIXAR.ToString & "') " & vbCrLf

            If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "05272759" Then
                sql &= "  and (prd.Grupo = '30101') " & vbCrLf
                'ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) <> "04854422" Then
                '    sql &= "  and (prd.Nome like '%SOJA%' or prd.Nome like '%MILHO%' or prd.Nome like 'LECITINA%' or prd.Grupo = '40201') " & vbCrLf
            End If

            If ddlEmpresa.SelectedIndex > 0 Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                sql &= "  and p.Empresa_id = '" & Empresa(0) & "' " & vbCrLf
            End If

            If txtCliente.Text.Length > 0 Then
                Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")

                If chkConsolidarCliente.Checked Then
                    sql &= " AND left(P.Cliente,8)    ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "  and p.Cliente    = '" & Cliente(0) & "'" & vbCrLf & _
                            "  and p.EndCliente = " & Cliente(1) & vbCrLf
                End If
            End If

            'Carrega sempre os bloqueados.
            If rdLiberados.Checked Then
                sql &= "  and isnull(p.PedidoBloqueado,0) = 0 " & vbCrLf
            ElseIf Not rdTodos.Checked Then
                sql &= "  and isnull(p.PedidoBloqueado,0) = 1 " & vbCrLf
            End If

            If txtConsultaPedido.Text.Length > 0 Then
                sql &= " and p.Pedido_id = '" & Trim(txtConsultaPedido.Text) & "' " & vbCrLf
            ElseIf chkPeriodo.Checked Then
                sql &= " and p.Datapedido between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
            End If

            If ddlUsuarioLiberacao.SelectedIndex > 0 Then
                sql &= " AND P.UsuarioLiberacao = '" & ddlUsuarioLiberacao.SelectedValue & "'" & vbCrLf
            End If

            If rdTrocaSim.Checked Then
                sql &= " AND ISNULL(P.Troca,0) = 1" & vbCrLf
            ElseIf rdTrocaNao.Checked Then
                sql &= " AND ISNULL(P.Troca,0) = 0" & vbCrLf
            End If

            sql &= "  and not  exists(select 1" & vbCrLf & _
                   "                    from notasfiscais nf " & vbCrLf & _
                   "                   where nf.empresa_id = p.empresa_id" & vbCrLf & _
                   "                     and nf.endempresa_id = p.endempresa_id" & vbCrLf & _
                   "                     and nf.pedido        = p.pedido_id" & vbCrLf & _
                   "                     and isnull(nf.nfg,0) = 1)" & vbCrLf

            If Quimica Then
                sql &= "group by p.Empresa_id, p.EndEmpresa_Id, cliEmp.Nome, cliEmp.Cidade, p.Pedido_Id, p.Cliente, p.EndCliente, cli.Nome," & vbCrLf & _
                       "p.DataPedido, p.Moeda, cope.Classe_id, prd.Grupo, gp.Descricao, isnull(p.PedidoBloqueado, 0), P.PedidoEfetivo" & vbCrLf
            End If

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
            dtPedidos.Columns.Add("Classe", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Moeda", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Produto", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("Unitario", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("Valor", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("UsuarioLiberacao", Type.GetType("System.String"))
            dtPedidos.Columns.Add("UsuarioLiberacaoData", Type.GetType("System.DateTime"))

            For Each dr As DataRow In ds.Tables(0).Rows
                Dim drPedido As DataRow = dtPedidos.NewRow()
                drPedido("Empresa") = dr("Empresa_Id") & "-" & Left(dr("Nome"), 20) & "/" & dr("Cidade")
                drPedido("Cliente") = dr("Cliente") & "-" & dr("EndCliente")
                drPedido("NomeCliente") = dr("NomeCliente")
                drPedido("Pedido") = dr("Pedido_Id")
                drPedido("PedidoEfetivo") = dr("PedidoEfetivo")
                drPedido("Data") = dr("DataPedido")
                drPedido("Classe") = dr("Classe_id")
                If dr("Moeda") = 1 Then
                    drPedido("Moeda") = "R$"
                Else
                    drPedido("Moeda") = "U$"
                End If
                drPedido("Produto") = dr("Produto_Id") & "-" & dr("NomeProduto")
                drPedido("Quantidade") = dr("Quantidade")
                drPedido("Unitario") = dr("Unitario")
                drPedido("Valor") = dr("Total")
                drPedido("UsuarioLiberacao") = dr("UsuarioLiberacao")
                drPedido("UsuarioLiberacaoData") = dr("UsuarioLiberacaoData")
                dtPedidos.Rows.Add(drPedido)
            Next

            gridConsulta.DataSource = dtPedidos
            gridConsulta.DataBind()
            Dim i As Integer = 0
            While i < gridConsulta.Rows.Count
                If ds.Tables(0).Rows(i).Item("PedidoBloqueado") = "True" Then
                    CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ImageUrl = "~/images/erro.jpg"
                    CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ToolTip = "Fechado"
                Else
                    CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ImageUrl = "~/images/certo.jpg"
                    CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ToolTip = "Aberto"
                End If
                i += 1
            End While
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

    Protected Sub lnkSelecionar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnkPedido As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnkPedido.NamingContainer, GridViewRow)
            Session("RowIndex" & HID.Value) = row.RowIndex
            gridConsulta.SelectedIndex = row.RowIndex

            Dim strPedido() As String = gridConsulta.Rows(row.RowIndex).Cells(1).Text.Split("-")

            objPedido = New [Lib].Negocio.Pedido(strPedido(0), 0, gridConsulta.Rows(row.RowIndex).Cells(4).Text)

            Limpar(False)

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
            txtOperacao.Text = objPedido.CodigoOperacao & "-" & objPedido.CodigoSubOperacao & " " & objPedido.SubOperacao.Descricao
            txtMoeda.Text = objPedido.Moeda.Descricao
            txtIndiceFixado.Text = objPedido.IndiceFixado


            txtProduto.Text = objPedido.Itens(0).CodigoProduto & "-" & objPedido.Itens(0).Produto.Nome
            txtQuantidade.Text = objPedido.Itens(0).QuantidadePedidoFaturamento

            If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                txtUnitario.Text = objPedido.Itens(0).UnitarioMedioFaturamento.ToString("N10")
                txtValor.Text = objPedido.Itens(0).PedidoValor.ToString("N2")

                'Como é moeda oficial os campos definidos para outras moedas serão sempre zero.
                txtUnitarioMoeda.Text = 0
                txtValorMoeda.Text = 0
            Else
                txtUnitarioMoeda.Text = objPedido.Itens(0).UnitarioMedioFaturamento.ToString("N10")
                txtValorMoeda.Text = objPedido.Itens(0).PedidoValor.ToString("N2")

                txtUnitario.Text = Funcoes.ConverteMoeda(objPedido.Itens(0).UnitarioMedioFaturamento, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                txtValor.Text = Funcoes.ConverteMoeda(objPedido.Itens(0).PedidoValor, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 2)
            End If

            Dim Parametros As New OperacaoXEstado
            Parametros.Empresa = Left(objPedido.CodigoEmpresa, 8)
            Parametros.CodigoGrupoProduto = objPedido.Itens(0).Produto.CodigoGrupo
            Parametros.CodigoProduto = objPedido.Itens(0).CodigoProduto
            Parametros.CodigoOperacao = objPedido.CodigoOperacao
            Parametros.CodigoSubOperacao = objPedido.CodigoSubOperacao
            Parametros.EstadoOrigem = objPedido.Empresa.CodigoEstado
            Parametros.EstadoDestino = objPedido.Cliente.CodigoEstado
            Dim OxE = New OperacaoXEstado(Parametros)


            If OxE.Encargos.Count > 0 Then
                Dim dtEncargoItem As New DataTable("EncargoItem")
                dtEncargoItem.Columns.Add("Codigo", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("SituacaoTributaria", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("SituacaoTributariaPISCOFINS", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("SituacaoTributariaIPI", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("Base", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("Percentual", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("PercentualExibicao", Type.GetType("System.String"))
                dtEncargoItem.Columns.Add("Valor", Type.GetType("System.Decimal"))
                dtEncargoItem.Columns.Add("Sinal", Type.GetType("System.String"))

                For Each EncOP In OxE.Encargos
                    If EncOP.CodigoEncargo = "PRODUTO" Then
                        Dim objCFOP As New [Lib].Negocio.CFOP(EncOP.OperacaoEstado.CodigoFiscal)
                        txtNaturezaDaOperacao.Text = EncOP.OperacaoEstado.CodigoFiscal & "-" & objCFOP.Descricao
                    End If

                    For Each EncPed As [Lib].Negocio.PedidoXEncargo In objPedido.Itens(0).Encargos
                        If EncPed.CodigoEncargo = EncOP.CodigoEncargo Then
                            Dim drRow As DataRow = dtEncargoItem.NewRow()
                            drRow("Codigo") = EncPed.CodigoEncargo
                            drRow("Base") = EncPed.BaseOficial
                            drRow("Percentual") = EncPed.Percentual
                            drRow("Valor") = EncPed.ValorOficial
                            Dim objSituacaoTributaria As New [Lib].Negocio.SituacaoTributaria(EncOP.OperacaoEstado.CodigoSTICMS)
                            drRow("SituacaoTributaria") = EncOP.OperacaoEstado.CodigoSTICMS & "-" & objSituacaoTributaria.Descricao
                            Dim objSituacaoTributariaPISCOFINS As New [Lib].Negocio.SituacaoTributariaPISCOFINS(EncOP.OperacaoEstado.CodigoSTPISCOFINS)
                            drRow("SituacaoTributariaPISCOFINS") = EncOP.OperacaoEstado.CodigoSTPISCOFINS & "-" & objSituacaoTributariaPISCOFINS.Descricao
                            Dim objSituacaoTributariaIPI As New [Lib].Negocio.SituacaoTributariaIPI(EncOP.OperacaoEstado.CodigoSTIPI)
                            drRow("SituacaoTributariaIPI") = EncOP.OperacaoEstado.CodigoSTIPI & "-" & objSituacaoTributariaIPI.Descricao
                            drRow("PercentualExibicao") = EncOP.AliquotaExibicao
                            drRow("Sinal") = EncOP.Sinal
                            dtEncargoItem.Rows.Add(drRow)
                        End If
                    Next
                Next

                gridEncargos.Parent.Visible = True
                gridEncargos.DataSource = dtEncargoItem
                gridEncargos.DataBind()
            End If

            'txtCifFob.Text = IIf(objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.CIF, " (CIF)", " (FOB)")
            If objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.CIF Then
                txtCifFob.Text = " (CIF)"
            ElseIf objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.FOB Then
                txtCifFob.Text = " (FOB)"
            ElseIf objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.TER Then
                txtCifFob.Text = " (TER)"
            Else
                txtCifFob.Text = " (NENHUM)"
            End If

            If objPedido.Transportadores.Count > 0 Then
                gridTransportes.Parent.Visible = True
                gridTransportes.DataSource = objPedido.Transportadores.ToArray
                gridTransportes.DataBind()
            End If

            If objPedido.Representantes.Count > 0 Then
                gridRepresentantes.Parent.Visible = True
                gridRepresentantes.DataSource = objPedido.Representantes.ToArray
                gridRepresentantes.DataBind()
            End If

            If Not FinanceiroNovo Then ddlMomentoFinanceiro.SelectedValue = objPedido.MomentoFinanceiro

            gridFinanceiro.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
            gridFinanceiro.DataBind()

            If objPedido.SubOperacao.Financeiro AndAlso objPedido.Vencimentos.Count > 0 Then
                pnlEntrega.Visible = True

                lstCondicoes.SelectedIndex = lstCondicoes.Items.IndexOf(lstCondicoes.Items.FindByValue(objPedido.CondicaoPagamento.Codigo.ToString()))

                If objPedido.CondicaoPagamentoDaEntrega Is Nothing Then
                    objPedido.CodigoCondicaoPagamentoDaEntrega = objPedido.CodigoCondicaoPagamento
                End If

                lstCondicoesPgtoEntrega.SelectedIndex = lstCondicoesPgtoEntrega.Items.IndexOf(lstCondicoesPgtoEntrega.Items.FindByValue(objPedido.CondicaoPagamentoDaEntrega.Codigo.ToString()))

                txtQuotaDeEntrega.Text = objPedido.QuotaEntrega
                ddlPeriodicidadeEntrega.SelectedValue = objPedido.PeriodicidadeEntrega
                If objPedido.PedidoBloqueado Then
                    btnAjustarEntrega.Enabled = False
                    ddlMomentoFinanceiro.Enabled = True
                    lstCondicoes.Enabled = True
                Else
                    ddlMomentoFinanceiro.Enabled = False
                    lstCondicoes.Enabled = False
                End If

                Dim objVencimentos As New Hashtable
                Dim k As Integer
                Dim temBaixa As Boolean = False
                For k = 0 To objPedido.Vencimentos.Count - 1
                    objVencimentos.Add(k, objPedido.Vencimentos(k).Codigo)

                    If objPedido.Vencimentos(k).Provisao = [Lib].Negocio.eProvisao.Baixa OrElse Not String.IsNullOrWhiteSpace(objPedido.Vencimentos(k).ContratoBancario) Then temBaixa = True
                Next
                Session("objPedVencimentos" & HID.Value) = objVencimentos

                If temBaixa Then
                    ddlMomentoFinanceiro.Enabled = False
                    lstCondicoes.Enabled = False
                    lstCondicoesPgtoEntrega.Enabled = False
                    txtQuotaDeEntrega.Enabled = False
                    ddlPeriodicidadeEntrega.Enabled = False
                    btnAjustarEntrega.Enabled = False
                End If
            Else
                lstCondicoesPgtoEntrega.SelectedIndex = 0
                txtQuotaDeEntrega.Text = String.Empty
                ddlPeriodicidadeEntrega.SelectedIndex = 0
                pnlEntrega.Visible = False
            End If

            If objPedido.PedidoBloqueado Then
                btnLiberar.BackColor = Drawing.Color.Red
                btnLiberar.Text = "Bloqueado"
            Else
                btnLiberar.BackColor = Drawing.Color.Green
                btnLiberar.Text = "Liberado"
            End If

            For Each item As [Lib].Negocio.PedidoXItem In objPedido.Itens
                If item.TemNota Then
                    ddlMomentoFinanceiro.Enabled = False
                    lstCondicoes.Enabled = False
                    lstCondicoesPgtoEntrega.Enabled = False
                    txtQuotaDeEntrega.Enabled = False
                    ddlPeriodicidadeEntrega.Enabled = False
                    btnAjustarEntrega.Enabled = False
                End If
            Next

            SalvarPedido()
            TabContainer1.ActiveTabIndex = 1
            'btnConsultar_Click(btnConsultar, Nothing)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnLiberar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLiberar.Click
        If Funcoes.VerificaPermissao("LiberacaoDePedido", "LIBERAR") Then
            RecuperarPedido()

            If objPedido.PedidoBloqueado Then
                btnLiberar.BackColor = Drawing.Color.Green
                btnLiberar.Text = "Liberado"
                objPedido.PedidoBloqueado = False
            Else
                btnLiberar.BackColor = Drawing.Color.Red
                btnLiberar.Text = "Bloqueado"
                objPedido.PedidoBloqueado = True
            End If

            If objPedido.AbrirFecharPedido(ajustarFinanceiro.Value) Then
                SalvarPedido()

                Dim i As Integer = 0
                While i < gridConsulta.Rows.Count
                    If gridConsulta.Rows(i).Cells(4).Text = objPedido.Codigo.ToString() Then
                        If objPedido.PedidoBloqueado Then
                            CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ImageUrl = "~/images/erro.jpg"
                            CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ToolTip = "Fechado"
                        Else
                            CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ImageUrl = "~/images/certo.jpg"
                            CType(gridConsulta.Rows(i).FindControl("imgStatus"), ImageButton).ToolTip = "Aberto"
                        End If
                    End If
                    i += 1
                End While

                If objPedido.PedidoBloqueado Then
                    MsgBox(Me.Page, "Pedido bloqueado com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, "Pedido liberado com Sucesso.", eTitulo.Sucess)
                End If
                Limpar(False)
                lnkConsultar_Click(lnkConsultar, Nothing)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para liberar pedido.")
        End If
    End Sub

    Protected Sub lstCondicoesPgtoEntrega_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If lstCondicoesPgtoEntrega.SelectedIndex > 0 Then
                RecuperarPedido()
                objPedido.CodigoCondicaoPagamentoDaEntrega = lstCondicoesPgtoEntrega.SelectedValue
                SalvarPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtQuotaDeEntrega_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(txtQuotaDeEntrega.Text) Then
                RecuperarPedido()
                If CDec(txtQuotaDeEntrega.Text) > CDec(txtQuantidade.Text) Then
                    txtQuotaDeEntrega.Text = txtQuantidade.Text
                    MsgBox(Me.Page, "Quota da entrega não pode ser maior que a quantidade disponível no pedido.")
                End If
                objPedido.QuotaEntrega = txtQuotaDeEntrega.Text
                SalvarPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlPeriodicidadeEntrega_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperarPedido()
            objPedido.PeriodicidadeEntrega = ddlPeriodicidadeEntrega.SelectedValue
            SalvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAjustarEntrega_Click(sender As Object, e As EventArgs) Handles btnAjustarEntrega.Click
        Try
            RecuperarPedido()
            If lstCondicoesPgtoEntrega.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Condições da entrega não foi selecionada.")
            ElseIf objPedido.QuotaEntrega > CDec(txtQuantidade.Text) Then
                txtQuotaDeEntrega.Text = txtQuantidade.Text
                MsgBox(Me.Page, "Quota da entrega não pode ser maior que a quantidade disponível no pedido.")
            Else
                If Funcoes.VerificaPermissao("LiberacaoDePedido", "LIBERAR") Then
                    RecuperarPedido()

                    objPedido.CodigoCondicaoPagamentoDaEntrega = lstCondicoesPgtoEntrega.SelectedValue
                    objPedido.QuotaEntrega = txtQuotaDeEntrega.Text
                    objPedido.PeriodicidadeEntrega = ddlPeriodicidadeEntrega.SelectedValue

                    If objPedido.AbrirFecharPedido(ajustarFinanceiro.Value) Then
                        MsgBox(Me.Page, "Pedido atualizado com Sucesso.", eTitulo.Sucess)

                        Limpar(False)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para ajustar pedido.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataEntrega_TextChanged(sender As Object, e As EventArgs) Handles txtDataEntrega.TextChanged
        Try
            If Not IsDate(txtDataEntrega.Text) Then
                MsgBox(Me.Page, "Data de entrada inválida.")
                txtDataEntrega.Text = Format(Today, "dd/MM/yyyy")
                Exit Sub
            End If

            RecuperarPedido()

            objPedido.DataEntregaInicial = CDate(txtDataEntrega.Text)

            SalvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMomentoFinanceiro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlMomentoFinanceiro.SelectedIndexChanged
        Try
            RecuperarPedido()

            If ddlMomentoFinanceiro.SelectedValue <> objPedido.MomentoFinanceiro Then
                objPedido.Vencimentos.Clear()
                lstCondicoes.SelectedIndex = 0
                gridFinanceiro.DataSource = Nothing
                gridFinanceiro.DataBind()
                ajustarFinanceiro.Value = True
                objPedido.CodigoCondicaoPagamento = 0
                LimparCondicoes(True)
            End If

            objPedido.MomentoFinanceiro = ddlMomentoFinanceiro.SelectedValue

            SalvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lstCondicoes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCondicoes.SelectedIndexChanged
        Try
            RecuperarPedido()

            If lstCondicoes.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Condição de pagamento não foi selecionada.")
            Else
                txtDataVencimento.Text = String.Empty
                txtValorVencimento.Text = String.Empty

                txtDataVencimento.Enabled = False
                txtValorVencimento.Enabled = False
                btnOkVencimento.Enabled = False

                ajustarFinanceiro.Value = True
                objPedido.CodigoCondicaoPagamento = lstCondicoes.SelectedValue

                If objPedido.SubOperacao.PrecoFixo Then
                    For Each objItem As [Lib].Negocio.PedidoXItem In objPedido.Itens
                        objItem.Fixacoes(0).CondicaoPagamento = objPedido.CondicaoPagamento.Codigo
                    Next
                End If
                CalcularParcelamento(False)
            End If

            SalvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridFinanceiro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridFinanceiro.SelectedIndexChanged
        Try
            If gridConsulta.SelectedRow.Cells(2).Text = "1" Then
                MsgBox(Me.Page, "Título baixado não pode ser alterado.")
            Else
                RecuperarPedido()

                For Each item As [Lib].Negocio.PedidoXItem In objPedido.Itens
                    If item.TemNota Then
                        MsgBox(Me.Page, "Pedido com Nota Fiscal não pode ser ajustado vencimentos por aqui, use o Financeiro.")
                        Exit Sub
                    End If
                Next

                txtDataVencimento.Enabled = True
                txtDataVencimento.Text = objPedido.Vencimentos(gridFinanceiro.SelectedIndex).Vencimento.ToString("dd/MM/yyyy")

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.MoedaEstrangeira Then
                    txtValorVencimento.Text = objPedido.Vencimentos(gridFinanceiro.SelectedIndex).ValorLiquidoMoeda.ToString("N2")
                Else
                    txtValorVencimento.Text = objPedido.Vencimentos(gridFinanceiro.SelectedIndex).ValorLiquidoOficial.ToString("N2")
                End If
                If gridFinanceiro.Rows.Count = 1 Or gridFinanceiro.SelectedIndex + 1 = gridFinanceiro.Rows.Count Then
                    txtValorVencimento.Enabled = False
                Else
                    txtValorVencimento.Enabled = True
                End If

                btnOkVencimento.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CalcularParcelamento()
        objPedido.Vencimentos.CriarParcelamento(False, CType(Session("objPedVencimentos" & HID.Value), Hashtable))
        gridFinanceiro.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
        gridFinanceiro.DataBind()

        objPedido.Vencimentos.TotalDasParcelas()
    End Sub

    Private Sub CalcularParcelamento(ByVal Manter As Boolean)
        objPedido.Vencimentos.CriarParcelamento(Manter, CType(Session("objPedVencimentos" & HID.Value), Hashtable))
        gridFinanceiro.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
        gridFinanceiro.DataBind()

        objPedido.Vencimentos.TotalDasParcelas()
    End Sub

    Private Sub LimparCondicoes(ByVal LimparGrade As Boolean)
        lstCondicoes.SelectedIndex = 0

        If LimparGrade Then
            gridFinanceiro.DataSource = Nothing
            gridFinanceiro.DataBind()
        End If
    End Sub

    Protected Sub btnOkVencimento_Click(sender As Object, e As EventArgs) Handles btnOkVencimento.Click
        Try
            RecuperarPedido()

            'Dim strData As String() = txtDataVencimento.Text.Split("/")
            objPedido.Vencimentos(gridFinanceiro.SelectedIndex).DataProrrogacao = Funcoes.ValidaDataUtil(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, CDate(txtDataVencimento.Text))

            Dim msg As String = String.Empty

            If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.MoedaEstrangeira Then
                'objPedido.Vencimentos.ModificarParcelaDolar(gridFinanceiro.SelectedIndex, _
                '                                       New DateTime(Convert.ToInt32(strData(2)), Convert.ToInt32(strData(1)), Convert.ToInt32(strData(0))), _
                '                                       Convert.ToDecimal(txtValorVencimento.Text))
                If Not objPedido.Vencimentos(gridFinanceiro.SelectedIndex).ValorDocumentoMoeda = CDec(txtValorVencimento.Text) Then msg = objPedido.Vencimentos.AjustaParcelas(gridFinanceiro.SelectedIndex, objPedido.Vencimentos(gridFinanceiro.SelectedIndex).ValorDocumentoMoeda, CDec(txtValorVencimento.Text))
            Else
                'objPedido.Vencimentos.ModificarParcela(gridFinanceiro.SelectedIndex, _
                '                                       New DateTime(Convert.ToInt32(strData(2)), Convert.ToInt32(strData(1)), Convert.ToInt32(strData(0))), _
                '                                       Convert.ToDecimal(txtValorVencimento.Text))
                If Not objPedido.Vencimentos(gridFinanceiro.SelectedIndex).ValorDocumentoOficial = CDec(txtValorVencimento.Text) Then msg = objPedido.Vencimentos.AjustaParcelas(gridFinanceiro.SelectedIndex, objPedido.Vencimentos(gridFinanceiro.SelectedIndex).ValorDocumentoOficial, CDec(txtValorVencimento.Text))
            End If

            gridFinanceiro.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
            gridFinanceiro.DataBind()

            txtDataVencimento.Text = String.Empty
            txtValorVencimento.Text = String.Empty

            txtDataVencimento.Enabled = False
            txtValorVencimento.Enabled = False
            btnOkVencimento.Enabled = False

            ajustarFinanceiro.Value = True

            SalvarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.ToString)
        End Try
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
            Funcoes.Ajuda(Me.Page, "LiberacaoDePedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class
