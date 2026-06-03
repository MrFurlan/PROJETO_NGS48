Imports NGS.Lib.Negocio

Public Class RelatorioDeAdiantamentos
    Inherits BasePage

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeAdiantamentos", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio)
                    ddl.Carregar(lstConta, CarregarDDL.Tabela.PlanoDeContas, "isnull(adiantamento,0) = 1 and len(conta_id) in (7,9)", False)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
                    txtDtBase.Text = DateTime.Now().ToString("dd/MM/yyyy")

                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaCliente_Click(sender As Object, e As EventArgs) Handles btnConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objCliente" & HID.Value.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            If validaCampos() Then
                Dim ds As DataSet = getDataSet()
                Dim param As New Dictionary(Of String, Object)
                getParametesConsulta(param)

                If chkEmpresa.Checked Then
                    Funcoes.BindReport(Me.Page, ds, "Cr_Adiantamento", eExportType.PDF, param)
                Else
                    Funcoes.BindReport(Me.Page, ds, "Cr_AdiantamentoCliente", eExportType.PDF, param)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If validaCampos() Then
                Dim ds As DataSet = getDataSet()
                Dim param As New Dictionary(Of String, Object)
                getParametesConsulta(param)

                If chkEmpresa.Checked Then
                    Funcoes.BindReport(Me.Page, ds, "Cr_Adiantamento", eExportType.ExcelCrystalDados, param)
                Else
                    Funcoes.BindReport(Me.Page, ds, "Cr_AdiantamentoCliente", eExportType.ExcelCrystalDados, param)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeAdiantamentos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods"
    Private Sub LimparCampos()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtPedido.Text = String.Empty
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function validaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a Empresa.")
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim cli As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
                Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(cli)
                txtCliente.Text = itemDeposito.Text
                txtCodigoCliente.Value = itemDeposito.Value
                Session.Remove("objCliente" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "Select cast(m.moeda_id as varchar) + '-' + m.descricao as moeda, a.Empresa_Id, a.EndEmpresa_Id, emp.Nome + ' - '  + emp.Cidade + ' / ' + emp.Estado + ' - ' + dbo.FormatarCpfCnpj(a.Empresa_Id) as DescEmpresa,                                            " & vbCrLf & _
                            "       a.Cliente_Id, a.EndCliente_Id, cli.Nome + ' - '  + cli.Cidade + ' / ' + cli.Estado  + ' - ' + dbo.FormatarCpfCnpj(a.Cliente_Id) + '/' + cast(a.EndCliente_Id as varchar) as DescCliente,  " & vbCrLf & _
                            "	    pc.Conta_Id, pc.Titulo as DescricaoConta, A.Adiantamento_Id, a.RegistroPedido as Pedido, a.Titulo, t.Movimento, a.Vencimento, a.ValorOficial, a.ValorMoeda, isnull(axb.TotalBaixado, 0.00) as TotalBaixado" & vbCrLf & _
                            "  into #Adiantamentos" & vbCrLf & _
                            "  from vw_Adiantamento a" & vbCrLf & _
                            "  Left Join (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Adiantamento_Id, sum(ValorOficial) as TotalBaixado" & vbCrLf & _
                            "               from vw_AdiantamentosxBaixas" & vbCrLf & _
                            "              Where DataBaixa <= '" & txtDtBase.Text.ToSqlDate & "'" & vbCrLf & _
                            "			   group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Adiantamento_Id" & vbCrLf & _
                            "			 ) as axb" & vbCrLf & _
                            "    on axb.Empresa_Id      = a.Empresa_Id" & vbCrLf & _
                            "   and axb.EndEmpresa_Id   = a.EndEmpresa_Id" & vbCrLf & _
                            "   and axb.Cliente_Id      = a.Cliente_Id" & vbCrLf & _
                            "   and axb.EndCliente_Id   = a.EndCliente_Id" & vbCrLf & _
                            "   and axb.Adiantamento_Id = a.Adiantamento_Id" & vbCrLf & _
                            " Inner Join (select registro_id, moeda, carteira, baixa as movimento, situacao, Baixa" & vbCrLf & _
                            "               from ContasAPagar" & vbCrLf & _
                            "              Union All" & vbCrLf & _
                            "			  select registro_id, moeda, carteira, baixa as movimento, situacao, Baixa" & vbCrLf & _
                            "			    from ContasAReceber" & vbCrLf & _
                            "		    ) t" & vbCrLf & _
                            "	 on a.Titulo = t.Registro_id" & vbCrLf & _
                            " Inner Join Moedas m" & vbCrLf & _
                            "    on m.moeda_id = t.moeda" & vbCrLf & _
                            " Inner join comprasxprodutos cart" & vbCrLf & _
                            "    on cart.produto_id   = t.carteira" & vbCrLf & _
                            " Inner Join PlanoDeContas pc" & vbCrLf & _
                            "    on pc.Conta_Id = cart.ContaClientes" & vbCrLf & _
                            " Inner Join Clientes emp" & vbCrLf & _
                            "    on emp.cliente_Id  = a.Empresa_Id" & vbCrLf & _
                            "   and emp.Endereco_Id = a.EndEmpresa_Id" & vbCrLf & _
                            " Inner Join Clientes cli" & vbCrLf & _
                            "    on cli.cliente_Id  = a.cliente_Id" & vbCrLf & _
                            "   and cli.Endereco_Id = a.EndCliente_Id" & vbCrLf & _
                            " Where t.situacao = 1 " & IIf(chkApenasComSaldo.Checked, " and a.ValorOficial <> isnull(axb.TotalBaixado, 0.00)", "") & vbCrLf & _
                            "   and t.Baixa <= '" & txtDtBase.Text.ToSqlDate & "'" & vbCrLf

        If lstConta.GetSelectedValues().Count > 0 Then sql &= "      and Pc.Conta_id in (" & String.Join(",", lstConta.GetSelectedValues().ToArray) & ")" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            If chkConsEmpresa.Checked Then
                sql &= "and left(a.Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= "and a.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            If chkConsCliente.Checked Then
                sql &= "      and left(a.Cliente_Id, 8)    = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= "      and a.Cliente_Id    = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                       "      and a.EndCliente_Id =  " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            End If
        End If

        If lstConta.GetSelectedValues().Count = 0 Then
            If rbFornecedor.Checked Then
                sql &= "      and left(pc.Conta_Id, 1) = 1" ' Conta Fornecedor
            Else
                sql &= "      and left(pc.Conta_Id, 1) = 2" ' Conta Cliente
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then sql &= "      and t.Pedido in (" & txtPedido.Text & ")" & vbCrLf
        sql &= " order by emp.Nome, cli.Nome, pc.Titulo, a.Adiantamento_Id, a.RegistroPedido, a.Titulo;" & vbCrLf


        sql &= "select moeda, Conta_Id, DescricaoConta, Adiantamento_Id, Pedido," & vbCrLf & _
               IIf(chkEmpresa.Checked, "Empresa_Id, EndEmpresa_Id, DescEmpresa,", "") & vbCrLf & _
               "       Cliente_Id, EndCliente_Id, DescCliente, Titulo, Movimento, Vencimento," & vbCrLf & _
               "	   ValorOficial, ValorMoeda, TotalBaixado" & vbCrLf & _
               "  From #Adiantamentos;" & vbCrLf

        sql &= "Select cast(m.moeda_id as varchar) + '-' + m.descricao as moeda," & vbCrLf & _
               IIf(chkEmpresa.Checked, "a.Empresa_Id, a.EndEmpresa_Id, a.DescEmpresa,", "") & vbCrLf & _
               "       a.Cliente_Id, a.EndCliente_Id, a.DescCliente," & vbCrLf & _
               "	   pc.Conta_Id, pc.Titulo as DescricaoConta, a.Adiantamento_Id, " & vbCrLf & _
               "       axb.Sequencia_Id, axb.RegistroPedido as Pedido, axb.Titulo, t.Movimento, axb.ValorOficial, axb.ValorMoeda, axb.DataBaixa " & vbCrLf & _
               "  from #Adiantamentos a" & vbCrLf & _
               " Inner Join vw_adiantamentosxbaixas axb" & vbCrLf & _
               "    on a.Empresa_Id      = axb.Empresa_Id" & vbCrLf & _
               "   and a.EndEmpresa_Id   = axb.EndEmpresa_Id" & vbCrLf & _
               "   and a.Cliente_Id      = axb.Cliente_Id" & vbCrLf & _
               "   and a.EndCliente_Id   = axb.EndCliente_Id" & vbCrLf & _
               "   and a.Adiantamento_Id = axb.adiantamento_Id" & vbCrLf & _
               "   and axb.DataBaixa    <= '" & txtDtBase.Text.ToSqlDate & "'" & vbCrLf & _
               " Inner Join (select registro_id, moeda, carteira, baixa as movimento" & vbCrLf & _
               "               from ContasAPagar" & vbCrLf & _
               "              Union All" & vbCrLf & _
               "			 select registro_id, moeda, carteira, baixa as movimento" & vbCrLf & _
               "			   from ContasAReceber" & vbCrLf & _
               "		    ) t" & vbCrLf & _
               "	on t.Registro_id  = axb.Titulo" & vbCrLf & _
               " Inner Join Moedas m" & vbCrLf & _
               "    on m.moeda_id = t.moeda" & vbCrLf & _
               " inner join comprasxprodutos cart" & vbCrLf & _
               "    on cart.produto_id   = t.carteira" & vbCrLf & _
               " Inner Join PlanoDeContas pc" & vbCrLf & _
               "    on pc.Conta_Id = cart.ContaClientes" & vbCrLf & _
               " Inner Join Clientes emp" & vbCrLf & _
               "    on emp.cliente_Id  = axb.Empresa_Id" & vbCrLf & _
               "   and emp.Endereco_Id = axb.EndEmpresa_Id" & vbCrLf & _
               " Inner Join Clientes cli" & vbCrLf & _
               "    on cli.cliente_Id  = axb.cliente_Id" & vbCrLf & _
               "   and cli.Endereco_Id = axb.EndCliente_Id" & vbCrLf

        If rbSintetico.Checked Then sql &= "where 1=2" & vbCrLf

        sql &= " order by emp.Nome, cli.Nome, pc.Titulo, axb.Adiantamento_Id, axb.Sequencia_Id, axb.RegistroPedido, axb.Titulo" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Adiantamento")
        ds.Tables(0).TableName = "Adiantamento"
        ds.Tables(1).TableName = "AdiantamentoXBaixa"

        Dim param As New Dictionary(Of String, Object)
        getParametesConsulta(param)
        Return ds
    End Function

    Private Sub getParametesConsulta(ByRef param As Dictionary(Of String, Object))
        param.Add("ParametroConsulta", "")
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then param("ParametroConsulta") = "Unidade: " & ddlUnidade.SelectedValue & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then param("ParametroConsulta") &= "Empresa: " & Funcoes.getDescricaoCliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1)) & vbCrLf
        If chkConsEmpresa.Checked Then param("ParametroConsulta") &= "Empresa Consolidada" & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then param("ParametroConsulta") &= "Cliente: " & Funcoes.getDescricaoCliente(txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1)) & vbCrLf
        param("ParametroConsulta") &= IIf(Not String.IsNullOrWhiteSpace(txtPedido.Text), "Pedido(s): " & txtPedido.Text, "") & vbCrLf
        param("ParametroConsulta") &= "Data Base: " & txtDtBase.Text
    End Sub

#End Region


End Class