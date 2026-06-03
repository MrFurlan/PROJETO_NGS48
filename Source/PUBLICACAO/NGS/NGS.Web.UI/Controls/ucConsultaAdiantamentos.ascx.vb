Imports System.Web
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaAdiantamentos
    Inherits BaseUserControl

    Dim sql As String = String.Empty
    Dim ds As DataSet
    Dim Parametros As Hashtable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub SessaoRecuperaParametros()
        Parametros = CType(Session("Parametros" + HID.Value), Hashtable)
    End Sub

    Private Sub SessaoSalvarDS()
        Session("ds" + HID.Value) = ds
    End Sub

    Private Sub SessaoRecuperaDS()
        ds = CType(Session("ds" + HID.Value), DataSet)
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView()
        SessaoRecuperaParametros()
        LitPedido.Text = "Pedido: " & Parametros("PedidoCliente")
        LitConta.Text = "Conta: " & Parametros("ContaContabilDescricao")
        LitMoeda.Text = "Moeda: " & Parametros("DescMoeda")
        If Parametros("Formulario") = "Financeiro" Then
            CargaAdiantamentos()
        End If
    End Sub

    Public Sub CargaAdiantamentos()
        ds = New DataSet
        sql = "SELECT A.Adiantamento_Id as NumeroDoAdto," & vbCrLf & _
              "       Titulos.Titulo," & vbCrLf & _
              "       isnull(A.RegistroPedido,0) as Pedido," & vbCrLf & _
              "       Titulos.Conta as Conta_Id," & vbCrLf & _
              "       convert(varchar,Titulos.Conta) + ' - ' +  PC.titulo AS Conta," & vbCrLf & _
              "       A.Safra, A.Vencimento, A.Taxa," & vbCrLf & _
              "       M.Moeda_Id as Moeda," & vbCrLf & _
              "	      M.Cifrao," & vbCrLf & _
              "	      Case" & vbCrLf & _
              "	        when M.Classificacao = 'O'" & vbCrLf & _
              "		       then A.ValorOficial" & vbCrLf & _
              "		       else A.ValorMoeda" & vbCrLf & _
              "       End as Adiantamento," & vbCrLf & _
              "	      Case" & vbCrLf & _
              "	        when M.Classificacao = 'O'" & vbCrLf & _
              "		       then isnull(AxB.VariacaoOficial,0)" & vbCrLf & _
              "		       else 0" & vbCrLf & _
              "       End as Juros," & vbCrLf & _
              "	      Case" & vbCrLf & _
              "	        when M.Classificacao = 'O'" & vbCrLf & _
              "		       then isnull(AxB.BaixaValorOficial,0)" & vbCrLf & _
              "		       else isnull(AxB.BaixaValorMoeda,0)" & vbCrLf & _
              "       End as Baixas," & vbCrLf & _
              "	      case when M.Classificacao = 'O' then A.ValorOficial else A.ValorMoeda End + case when M.Classificacao = 'O' then isnull(AxB.VariacaoOficial,0) else 0 End - case when M.Classificacao = 'O' then isnull(AxB.BaixaValorOficial,0) else isnull(AxB.BaixaValorMoeda,0) End as Saldo" & vbCrLf & _
              "  FROM vw_Adiantamento A " & vbCrLf & _
              "  LEFT JOIN (Select Empresa_Id," & vbCrLf & _
              "                    EndEmpresa_Id," & vbCrLf & _
              "				       Cliente_Id," & vbCrLf & _
              "				       EndCliente_Id," & vbCrLf & _
              "				       Adiantamento_Id," & vbCrLf & _
              "				       Sum(isnull(ValorOficial,0))    as BaixaValorOficial," & vbCrLf & _
              "				       Sum(isnull(ValorMoeda,0))      as BaixaValorMoeda," & vbCrLf & _
              "				       Sum(isnull(VariacaoOficial,0)) as VariacaoOficial" & vbCrLf & _
              "              from vw_AdiantamentosXBaixas" & vbCrLf & _
              "             Where Titulo <> " & Parametros("Titulo") & vbCrLf & _
              "               AND Empresa_Id    = '" & Parametros("Empresa") & "'" & vbCrLf & _
              "               AND EndEmpresa_Id =  " & Parametros("EndEmpresa") & vbCrLf & _
              "               AND Cliente_Id    = '" & Parametros("Cliente") & "'" & vbCrLf & _
              "               AND EndCliente_Id =  " & Parametros("EndCliente") & vbCrLf & _
              "             group by Empresa_Id," & vbCrLf & _
              "				         EndEmpresa_Id," & vbCrLf & _
              "				    	 Cliente_Id," & vbCrLf & _
              "				    	 EndCliente_Id," & vbCrLf & _
              "				    	 Adiantamento_Id" & vbCrLf & _
              "            ) AxB" & vbCrLf & _
              "    ON A.Empresa_ID      = AxB.Empresa_Id" & vbCrLf & _
              "   AND A.EndEmpresa_ID   = AxB.EndEmpresa_Id" & vbCrLf & _
              "   AND A.Cliente_ID      = AxB.Cliente_Id" & vbCrLf & _
              "   AND A.EndCliente_ID   = AxB.EndCliente_Id" & vbCrLf & _
              "   AND A.Adiantamento_Id = AxB.Adiantamento_Id" & vbCrLf

        If FinanceiroNovo Then
            sql &= "  Left join (Select t.Titulo_id as Titulo, Conta_id, t.moeda, convert(varchar,pc.Conta_Id) + ' - ' +  pc.titulo as Conta" & vbCrLf & _
                   "               from Titulos t" & vbCrLf & _
                   "			  inner join PlanoDeContas PC" & vbCrLf & _
                   "			     on t.ContaContabilCliFor = pc.Conta_Id" & vbCrLf & _
                   "            ) as Titulos" & vbCrLf
        Else
            sql &= "  Left join (Select cr.Registro_id as Titulo, Cart.ContaClientes as Conta, cr.Moeda" & vbCrLf & _
                   "               from ContasAReceber cr" & vbCrLf & _
                   "              inner join ComprasXProdutos Cart" & vbCrLf & _
                   "			     on cr.Carteira = Cart.Produto_Id" & vbCrLf & _
                   "             WHERE cr.Empresa    = '" & Parametros("Empresa") & "'" & vbCrLf & _
                   "               AND cr.EndEmpresa =  " & Parametros("EndEmpresa") & vbCrLf & _
                   "               AND cr.Cliente    = '" & Parametros("Cliente") & "'" & vbCrLf & _
                   "               AND cr.EndCliente =  " & Parametros("EndCliente") & vbCrLf & _
                   "			  union All" & vbCrLf & _
                   "			 Select cp.Registro_id as Titulo, Cart.ContaClientes as Conta, cp.Moeda" & vbCrLf & _
                   "               from ContasAPagar cp" & vbCrLf & _
                   "              inner join ComprasXProdutos Cart" & vbCrLf & _
                   "			     on cp.Carteira = Cart.Produto_Id" & vbCrLf & _
                   "             WHERE cp.Empresa    = '" & Parametros("Empresa") & "'" & vbCrLf & _
                   "               AND cp.EndEmpresa =  " & Parametros("EndEmpresa") & vbCrLf & _
                   "               AND cp.Cliente    = '" & Parametros("Cliente") & "'" & vbCrLf & _
                   "               AND cp.EndCliente =  " & Parametros("EndCliente") & vbCrLf & _
                   "			  union All" & vbCrLf & _
                   "			 Select 0 as Titulo, " & Parametros("ContaContabil") & " as Conta, " & Parametros("Moeda") & " as Moeda" & vbCrLf & _
                   "            ) as Titulos" & vbCrLf
        End If

        sql &= "    on A.Titulo = Titulos.Titulo" & vbCrLf & _
               "  Left Join Moedas M" & vbCrLf & _
               "    on Titulos.Moeda = m.Moeda_Id" & vbCrLf & _
               "  inner join PlanoDeContas PC" & vbCrLf & _
               "    on Titulos.Conta = PC.Conta_Id" & vbCrLf & _
               " WHERE A.Empresa_ID    = '" & Parametros("Empresa") & "'" & vbCrLf & _
               "   AND A.EndEmpresa_ID =  " & Parametros("EndEmpresa") & vbCrLf & _
               "   AND A.Cliente_ID    = '" & Parametros("Cliente") & "'" & vbCrLf & _
               "   AND A.EndCliente_ID =  " & Parametros("EndCliente") & vbCrLf & _
               "   AND Titulos.Conta   = '" & Parametros("ContaContabil") & "'" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Consulta")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            GridAdiantamentos.DataSource = ds
            GridAdiantamentos.DataBind()
        Else
            GridAdiantamentos.DataSource = Nothing
            GridAdiantamentos.DataBind()
        End If

        SessaoSalvarDS()
        rdTodos.Checked = True
        rdComSaldo.Checked = False
    End Sub

    Protected Sub GridAdiantamentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridAdiantamentos.SelectedIndexChanged
        SessaoRecuperaParametros()
        SessaoRecuperaDS()
        Dim Adiantamento As String = GridAdiantamentos.SelectedRow.Cells(1).Text()
        Dim Pedido As String = GridAdiantamentos.SelectedRow.Cells(3).Text()
        Dim Saldo As Decimal = CDec(GridAdiantamentos.SelectedRow.Cells(12).Text())

        If CInt(Parametros("Moeda")) <> CInt(ds.Tables(0).Select("NumeroDoAdto=" & Adiantamento)(0)(8)) Then
            MsgBox(Me.Page, "Somente Adiantamentos em Moeda " & Parametros("DescMoeda") & " podem ser Selecionados para esta operação.")
            Exit Sub
        End If

        If CInt(Parametros("ContaContabil")) <> CInt(GridAdiantamentos.SelectedRow.Cells(4).Text().Split("-")(0)) Then
            MsgBox(Me.Page, "Somente Adiantamentos com a conta contabil " & Parametros("ContaContabil") & " podem ser Selecionados para esta operação.")
            Exit Sub
        End If

        If IsNumeric(Pedido) AndAlso CInt(Pedido) > 0 AndAlso CInt(Parametros("Pedido")) <> CInt(Pedido) Then
            MsgBox(Me.Page, "As Baixas de Adiantamento devem ocorrer dentro do mesmo pedido")
            Exit Sub
        End If

        If Parametros("Formulario") = "Financeiro" Then
            Dim txtNumeroAdto As TextBox = CType(WebHelpers.FindControlRecursive(Me.Page, "txtNumeroAdto"), TextBox)
            Dim txtPedido As TextBox = CType(WebHelpers.FindControlRecursive(Me.Page, "txtPedido"), TextBox)
            Dim HDSaldoAdiantamento As TextBox = CType(WebHelpers.FindControlRecursive(Me.Page, "Saldo"), TextBox)

            If txtNumeroAdto IsNot Nothing Then
                txtNumeroAdto.Text = Adiantamento
            End If

            If txtPedido IsNot Nothing Then
                If Not IsNumeric(txtPedido.Text) OrElse CInt(txtPedido.Text) <= 0 Then
                    txtPedido.Text = IIf(IsNumeric(Pedido), Pedido, 0)
                End If
            End If

            If TypeOf Me.Page Is CompensacaoDeAdiantamentos Then
                If Saldo > 0 Then
                    CType(Me.Page, CompensacaoDeAdiantamentos).CarregarAdiantamento(ds, GridAdiantamentos.SelectedRow.RowIndex)
                    Popup.CloseDialog(Me.Page, "divConsultaAdiantamento")
                Else
                    MsgBox(Me.Page, "Somente Adiantamentos com saldo podem ser Selecionados para esta operação.")
                End If
            Else
                Popup.CloseDialog(Me.Page, "divConsultaAdiantamento")
            End If
        End If
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaAdiantamento")
    End Sub

    Protected Sub rdTodos_CheckedChanged(sender As Object, e As EventArgs) Handles rdTodos.CheckedChanged
        Try
            SessaoRecuperaDS()
            GridAdiantamentos.DataSource = ds
            GridAdiantamentos.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub rdComSaldo_CheckedChanged(sender As Object, e As EventArgs) Handles rdComSaldo.CheckedChanged
        SessaoRecuperaDS()
        Dim dv As New DataView(ds.Tables(0), "Saldo > 0", "Saldo", DataViewRowState.OriginalRows)

        GridAdiantamentos.DataSource = dv
        GridAdiantamentos.DataBind()
    End Sub

    Protected Sub rdSelecionaveis_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelecionaveis.CheckedChanged
        SessaoRecuperaParametros()
        SessaoRecuperaDS()

        Dim where As String = "Saldo > 0"
        where &= " and Conta_id = '" & Parametros("ContaContabil") & "'"
        where &= " and Moeda = " & Parametros("Moeda")
        If Parametros("Pedido") > 0 Then where &= " and (Pedido = " & Parametros("Pedido") & " or Pedido = 0)"

        Dim dv As New DataView(ds.Tables(0), where, "Saldo", DataViewRowState.OriginalRows)

        GridAdiantamentos.DataSource = dv
        GridAdiantamentos.DataBind()
    End Sub
End Class