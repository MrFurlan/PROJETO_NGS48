Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class ConferenciaFiscalFinanceiroPedido
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PosicaoDePedidos", "ACESSAR") Then
                    Me.setMenu(eModulo.Comercial)
                    ddl.Carregar(cmbUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(cmbUnidadeNegocio, ddlEmpresa)
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
                    ddl.Carregar(lstClasseOp, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1", False)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub



    Protected Sub btnConsultaCliente_Click(sender As Object, e As EventArgs) Handles btnConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If validaCampo() Then
                Dim ds As DataSet = getDataset()
                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then MsgBox(Me.Page, "Nenhum resultado encontrado.")
                gridPedido.DataSource = ds
                gridPedido.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConferenciaFiscalFinanceiroPedido")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExtrato_Click(sender As Object, e As EventArgs)
        Try
            Dim index As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
            Dim pedido As Integer = gridPedido.Rows(index).Cells(0).Text
            Extrato.Emitir(Me.Page, FinanceiroNovo, ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), "T", pedido)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Function validaCampo() As Boolean
        If String.IsNullOrWhiteSpace(cmbUnidadeNegocio.SelectedValue) AndAlso String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a Unidade de Negocio ou a empresa.")
            Return False
        ElseIf (String.IsNullOrWhiteSpace(txtData1.Text) OrElse String.IsNullOrWhiteSpace(txtData2.Text)) AndAlso String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            MsgBox(Me.Page, "Informe um período e/ou safra.")
            Return False
        End If
        Return True
    End Function

    Private Function getDataset() As DataSet
        Dim sql As String
        sql = "Select p.Pedido_Id as Pedido," & vbCrLf & _
              "       dbo.FormatarCpfCnpj(p.Cliente) + ' - ' + cast(p.EndCliente as varchar) + ' - ' + c.Nome + ' - ' + c.Cidade + '/' + c.Estado as Cliente,  " & vbCrLf & _
              "	      p.Safra," & vbCrLf & _
              "       case when isnull(P.Antecipada,0) = 1 then 'SIM' Else 'NAO' end Antecipada," & vbCrLf & _
              "       case when isnull(P.Recompra  ,0) = 1 then 'SIM' Else 'NAO' end Recompra," & vbCrLf & _
              "       case when isnull(P.Troca     ,0) = 1 then 'SIM' Else 'NAO' end Troca," & vbCrLf & _
              "       case when P.FinanceiroAberto = 1 then 'SIM' Else 'NAO' end FinanceiroAberto," & vbCrLf & _
              "       case when P.FiscalAberto     = 1 then 'SIM' Else 'NAO' end FiscalAberto" & vbCrLf & _
              "  from pedidos p" & vbCrLf & _
              " inner join Operacoes OP" & vbCrLf & _
              "    on OP.Operacao_id = p.operacao" & vbCrLf & _
              " Inner Join Clientes c" & vbCrLf & _
              "    on c.Cliente_Id     = p.Cliente" & vbCrLf & _
              "   and c.Endereco_Id    = p.EndCliente" & vbCrLf & _
              " where P.situacao in (1,4,50)" & vbCrLf

        If cmbUnidadeNegocio.SelectedIndex > 0 Then
            sql &= "  and p.UnidadeDeNegocio ='" & cmbUnidadeNegocio.SelectedValue & "'"
        End If

        If ddlEmpresa.SelectedIndex > 0 Then
            If chkConsolidarEmpresa.Checked Then
                sql &= "   and left(p.Empresa_Id,8) ='" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= "   and p.Empresa_Id     ='" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If
        End If

        sql &= "   and Exists(Select 1" & vbCrLf & _
               "                from PedidoxItem pxi" & vbCrLf & _
               "               Inner Join Produtos prd" & vbCrLf & _
               "                  on prd.Produto_Id = pxi.Produto_id" & vbCrLf & _
               "               Where pxi.Empresa_Id    = p.Empresa_Id" & vbCrLf & _
               "                 and pxi.EndEmpresa_Id = p.EndEmpresa_Id" & vbCrLf & _
               "                 and pxi.Pedido_Id     = p.Pedido_Id" & vbCrLf
        If ucSelecaoProduto.TemSelecionado() Then sql &= " AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("prd.Grupo", "prd.Produto_Id", "")(0)
        sql &= "              )" & vbCrLf

        If chkFinanceiroAberto.Checked And chkFiscalAberto.Checked Then
            sql &= "   and (p.financeiroaberto = 1 or p.FiscalAberto = 1)" & vbCrLf
        ElseIf chkFinanceiroAberto.Checked Then
            sql &= "   and p.financeiroaberto = 1" & vbCrLf
        ElseIf chkFiscalAberto.Checked Then
            sql &= "   and p.FiscalAberto     = 1" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            If chkConsolidarCliente.Checked Then
                sql &= "   and Left(p.Cliente,8)  = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'" & vbCrLf

            Else
                sql &= "   and p.Cliente          = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                       "   and p.EndCliente       =  " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            End If

        End If

        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            sql &= "   and p.Safra            = '" & ddlSafra.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not String.IsNullOrWhiteSpace(txtData2.Text) Then
            sql &= "   and p.DataPedido between '" & txtData1.Text & "' and '" & txtData2.Text & "'" & vbCrLf
        End If

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            sql &= "   AND op.classe in ('" & op & "')"
        End If

        Return Banco.ConsultaDataSet(sql, "Pedido")
    End Function

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(obj)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub LimparCampos()
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        ddlSafra.SelectedValue = String.Empty
        txtData1.Text = String.Empty
        txtData2.Text = String.Empty
        ucSelecaoProduto.Limpar()
        chkFinanceiroAberto.Checked = False
        chkFiscalAberto.Checked = False
        gridPedido.DataBind()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub


    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbUnidadeNegocio.SelectedIndexChanged
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidadeNegocio.SelectedValue, True)
    End Sub


End Class