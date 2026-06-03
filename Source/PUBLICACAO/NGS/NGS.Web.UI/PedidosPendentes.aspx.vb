Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PedidosPendentes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PedidosPendentes", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa)
                    ddl.Carregar(DdlSafra, CarregarDDL.Tabela.Safra)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        chkRelAnalitico.Checked = True
        chkRelResumoClienteProduto.Checked = False
        chkRelResumoCliente.Checked = False
        chkConsolidarEmpresasApresentacao.Checked = False
        chkConsolidarClientesApresentacao.Checked = False
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa)
        ChkConsolidarEmpresa.Checked = False
        DdlSafra.SelectedIndex = 0
        ddlClasse.SelectedIndex = 0
        chkPeriodo.Checked = False
        txtDataInicial.Text = String.Empty
        txtDataFinal.Text = String.Empty
        ChkConsolidarCliente.Checked = False
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf Not chkRelAnalitico.Checked AndAlso Not chkRelResumoClienteProduto.Checked AndAlso Not chkRelResumoCliente.Checked Then
            MsgBox(Me.Page, "Tipo de relatório não selecionado, por padrão será impresso como Analítico.")
            chkRelAnalitico.Checked = True
        End If
        Return True

    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Session("objClienteCxN" & HID.Value) IsNot Nothing Then
                Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteCxN" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteCxN" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCxN" & HID.Value, "txtNome")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PedidosPendentes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub chkPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles chkPeriodo.CheckedChanged
        Try
            pnlDataMovimento.Visible = chkPeriodo.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = ""
        Dim Cliente() As String = txtCodigoCliente.Value.Split("-")
        Dim Empresa() As String = DdlEmpresa.SelectedValue.Split("-")

        If chkConsolidarEmpresasApresentacao.Checked Then
            sql = "SELECT left(P.Empresa_id,8) as Empresa_id," & vbCrLf & _
                  "       0 as EndEmpresa_id," & vbCrLf & _
                  "       (select top 1 Nome from clientes where left(clientes.cliente_id,8) = left(P.Empresa_Id,8)) as  NomeEmpresa," & vbCrLf & _
                  "       '' as CidadeEmpresa," & vbCrLf & _
                  "       '' as EstadoEmpresa," & vbCrLf
        Else
            sql = "SELECT P.Empresa_id," & vbCrLf & _
                  "       P.EndEmpresa_id," & vbCrLf & _
                  "       E.Nome as NomeEmpresa," & vbCrLf & _
                  "       E.Cidade as CidadeEmpresa," & vbCrLf & _
                  "       E.Estado as EstadoEmpresa," & vbCrLf
        End If

        If chkConsolidarClientesApresentacao.Checked Then
            sql &= "       left(P.Cliente,8) as Cliente," & vbCrLf & _
                   "       0 as EndCliente," & vbCrLf & _
                   "       (select top 1 Nome from clientes where left(clientes.cliente_id,8) = left(P.Cliente,8)) as Nome," & vbCrLf & _
                   "       '' as Complemento," & vbCrLf & _
                   "       '' as Cidade," & vbCrLf & _
                   "       '' as Estado," & vbCrLf
        Else
            sql &= "       P.Cliente," & vbCrLf & _
                   "       P.EndCliente," & vbCrLf & _
                   "       c.Nome," & vbCrLf & _
                   "       C.Complemento," & vbCrLf & _
                   "       C.Cidade," & vbCrLf & _
                   "       C.Estado," & vbCrLf
        End If

        sql &= "       P.Pedido_Id," & vbCrLf & _
               "       P.Safra," & vbCrLf & _
               "       pxi.Quantidade AS QuantidadePedido," & vbCrLf & _
               "       pxi.unitarioOficial as UnitarioPedido," & vbCrLf & _
               "       pxi.TotalOficial AS ValorPedido," & vbCrLf & _
               "       SUM(CASE" & vbCrLf & _
               "             WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "               THEN ISNULL(NxI.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
               "               ELSE ISNULL(NxI.QuantidadeFiscal, 0)" & vbCrLf & _
               "           END) AS QuantidadeEntregue," & vbCrLf & _
               "       Case" & vbCrLf & _
               "         When SUM(CASE" & vbCrLf & _
               "					 WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "					   THEN ISNULL(NxI.Valor, 0) * - 1" & vbCrLf & _
               "					   ELSE ISNULL(NxI.Valor, 0)" & vbCrLf & _
               "				   END) = 0" & vbCrLf & _
               "		      Then 0" & vbCrLf & _
               "		      Else SUM(CASE" & vbCrLf & _
               "					  WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "					    THEN ISNULL(NxI.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
               "					    ELSE ISNULL(NxI.QuantidadeFiscal, 0)" & vbCrLf & _
               "				   END) / SUM(CASE" & vbCrLf & _
               "								WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "								  THEN ISNULL(NxI.Valor, 0) * - 1" & vbCrLf & _
               "								  ELSE ISNULL(NxI.Valor, 0)" & vbCrLf & _
               "							  END)" & vbCrLf & _
               "	      end UnitarioNota," & vbCrLf & _
               "       SUM(CASE" & vbCrLf & _
               "             WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "               THEN ISNULL(NxI.Valor, 0) * - 1" & vbCrLf & _
               "               ELSE ISNULL(NxI.Valor, 0)" & vbCrLf & _
               "           END) AS ValorNota," & vbCrLf & _
               "       Prd.Produto_Id," & vbCrLf & _
               "       Prd.Unidade," & vbCrLf & _
               "       Prd.Nome AS NomeDoProduto" & vbCrLf & _
               "  Into #Base" & vbCrLf & _
               "  FROM Pedidos P" & vbCrLf & _
               "  Left JOIN VW_PedidosXItensXFixacoes PxI" & vbCrLf & _
               "    ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
               "   AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
               "   AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
               " Inner Join Clientes C" & vbCrLf & _
               "    on C.Cliente_id  = P.Cliente" & vbCrLf & _
               "   and C.endereco_id = P.EndCliente" & vbCrLf & _
               " Inner Join Clientes E" & vbCrLf & _
               "    on E.Cliente_id  = P.Empresa_id" & vbCrLf & _
               "   and E.endereco_id = P.EndEmpresa_id" & vbCrLf & _
               " INNER JOIN SubOperacoes SOPD" & vbCrLf & _
               "    ON P.Operacao    = SOPD.Operacao_Id" & vbCrLf & _
               "   AND P.SubOperacao = SOPD.SubOperacoes_Id" & vbCrLf & _
               "  LEFT Join NotasFiscais NF" & vbCrLf & _
               "    ON NF.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
               "   AND NF.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
               "   AND NF.Pedido        = PxI.Pedido_Id" & vbCrLf & _
               "  LEFT JOIN NotasFiscaisXItens NxI" & vbCrLf & _
               "    ON NF.Empresa_Id      = NxI.Empresa_Id" & vbCrLf & _
               "   AND NF.EndEmpresa_Id   = NxI.EndEmpresa_Id" & vbCrLf & _
               "   AND NF.Cliente_id      = NxI.Cliente_Id" & vbCrLf & _
               "   AND NF.EndCliente_Id   = NxI.EndCliente_Id" & vbCrLf & _
               "   AND NF.EntradaSaida_id = NxI.EntradaSaida_id" & vbCrLf & _
               "   AND NF.Nota_id         = NxI.Nota_id" & vbCrLf & _
               "   AND NF.serie_id        = NxI.serie_id" & vbCrLf & _
               "   AND PxI.Produto_Id     = NxI.Produto_Id" & vbCrLf & _
               " INNER JOIN SubOperacoes SONF" & vbCrLf & _
               "    ON NxI.Operacao    = SONF.Operacao_Id" & vbCrLf & _
               "   AND NxI.SubOperacao = SONF.SubOperacoes_Id" & vbCrLf & _
               " Inner Join Produtos Prd" & vbCrLf & _
               "    ON prd.Produto_id = Pxi.Produto_Id" & vbCrLf & _
               " WHERE P.Situacao    =  1" & vbCrLf & _
               "   AND SOPD.CLASSE   = '" & ddlClasse.SelectedValue & "'" & vbCrLf & _
               "   AND SONF.CLASSE  <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf & _
               "   and P.FiscalAberto = 1" & vbCrLf

        If DdlSafra.SelectedIndex > 0 Then
            sql &= "   and P.Safra = '" & DdlSafra.SelectedValue & "'" & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Pxi.Produto_id")
            sql &= " AND " & retorno(0)
        End If

        If chkPeriodo.Checked Then
            sql &= "   and P.DataPedido BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If DdlEmpresa.SelectedIndex > 0 Then
            If ChkConsolidarEmpresa.Checked Then
                sql &= "   and left(P.empresa_id) = '" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
            Else
                sql &= "   and P.Empresa_id    ='" & Empresa(0) & "'" & vbCrLf & _
                       "   and P.EndEmpresa_id = " & Empresa(1) & vbCrLf
            End If
        End If

        If txtCliente.Text <> "" Then
            If ChkConsolidarCliente.Checked Then
                sql &= "   and Left(P.Cliente,8)    ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
            Else
                sql &= "   and P.Cliente    ='" & Cliente(0) & "'" & vbCrLf & _
                       "   and P.EndCliente = " & Cliente(1) & vbCrLf
            End If
        End If

        If chkConsolidarEmpresasApresentacao.Checked Then
            sql &= " GROUP BY left(P.Empresa_id,9)," & vbCrLf

        Else
            sql &= " GROUP BY P.Empresa_id," & vbCrLf & _
                   "          P.EndEmpresa_id," & vbCrLf & _
                   "          E.Nome," & vbCrLf & _
                   "          E.Cidade," & vbCrLf & _
                   "          E.Estado," & vbCrLf
        End If

        If chkConsolidarClientesApresentacao.Checked Then
            sql &= "          left(P.Cliente,8)," & vbCrLf
        Else
            sql &= "          P.Cliente," & vbCrLf & _
                   "          P.EndCliente," & vbCrLf & _
                   "          C.Nome," & vbCrLf & _
                   "          C.Complemento," & vbCrLf & _
                   "          C.Cidade," & vbCrLf & _
                   "          C.Estado," & vbCrLf
        End If

        sql &= "          P.Pedido_Id," & vbCrLf & _
               "          P.Safra," & vbCrLf & _
               "          Prd.Produto_Id," & vbCrLf & _
               "          Prd.Unidade," & vbCrLf & _
               "          Prd.Nome," & vbCrLf & _
               "          PxI.UnitarioOficial," & vbCrLf & _
               "          pxi.Quantidade," & vbCrLf & _
               "          pxi.unitarioOficial," & vbCrLf & _
               "          pxi.TotalOficial" & vbCrLf & _
               "  Having (" & vbCrLf & _
               "          pxi.Quantidade   <> SUM(CASE" & vbCrLf & _
               "	  						        WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "	  							      THEN ISNULL(NxI.QuantidadeFiscal, 0) * - 1" & vbCrLf & _
               "	  								  ELSE ISNULL(NxI.QuantidadeFiscal, 0)" & vbCrLf & _
               "	  							  END)" & vbCrLf & _
               "          or" & vbCrLf & _
               "          pxi.TotalOficial <> SUM(CASE" & vbCrLf & _
               "                                    WHEN SONF.Devolucao = 'S'" & vbCrLf & _
               "                                      THEN ISNULL(NxI.Valor, 0) * - 1" & vbCrLf & _
               "                                      ELSE ISNULL(NxI.Valor, 0)" & vbCrLf & _
               "                                  END)" & vbCrLf & _
               "          )" & vbCrLf & _
               "    order by p.safra" & vbCrLf

        sql &= "Select * from #Base" & vbCrLf

        '**************************************************************************
        '*******************   Cliente / Produto **********************************
        '**************************************************************************
        sql &= "select Cliente," & vbCrLf & _
               "       EndCliente," & vbCrLf & _
               "       Nome," & vbCrLf & _
               "       Complemento," & vbCrLf & _
               "       Cidade," & vbCrLf & _
               "       Estado," & vbCrLf & _
               "       Produto_Id," & vbCrLf & _
               "       Unidade," & vbCrLf & _
               "       NomeDoProduto," & vbCrLf & _
               "       sum(QuantidadePedido) as QuantidadePedido," & vbCrLf & _
               "       Case" & vbCrLf & _
               "         When sum(QuantidadePedido) = 0" & vbCrLf & _
               "           then 0.00" & vbCrLf & _
               "           else sum(ValorPedido) / sum(QuantidadePedido)" & vbCrLf & _
               "       end as UnitarioPedido," & vbCrLf & _
               "       sum(ValorPedido) as ValorPedido," & vbCrLf & _
               "       sum(QuantidadeEntregue) as QuantidadeEntregue," & vbCrLf & _
               "       Case" & vbCrLf & _
               "         When sum(QuantidadeEntregue) = 0" & vbCrLf & _
               "           then 0.00" & vbCrLf & _
               "           else  sum(ValorNota) / sum(QuantidadeEntregue)" & vbCrLf & _
               "       end as UnitarioNota," & vbCrLf & _
               "       sum(ValorNota) as ValorNota" & vbCrLf & _
               "  from #base" & vbCrLf & _
               " group by Cliente," & vbCrLf & _
               "          EndCliente," & vbCrLf & _
               "          Nome," & vbCrLf & _
               "          Complemento," & vbCrLf & _
               "          Cidade," & vbCrLf & _
               "          Estado," & vbCrLf & _
               "          Produto_Id," & vbCrLf & _
               "          Unidade," & vbCrLf & _
               "          NomeDoProduto" & vbCrLf

        '**************************************************************************
        '*******************   Cliente           **********************************
        '**************************************************************************
        sql &= "Select Cliente," & vbCrLf & _
               "       EndCliente," & vbCrLf & _
               "       Nome," & vbCrLf & _
               "       Complemento," & vbCrLf & _
               "       Cidade," & vbCrLf & _
               "       Estado," & vbCrLf & _
               "       sum(ValorPedido) as ValorPedido," & vbCrLf & _
               "       sum(ValorNota)   as ValorNota" & vbCrLf & _
               "  from #base" & vbCrLf & _
               " group by Cliente," & vbCrLf & _
               "          EndCliente," & vbCrLf & _
               "          Nome," & vbCrLf & _
               "          Complemento," & vbCrLf & _
               "          Cidade," & vbCrLf & _
               "          Estado" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "PedidosPendente")
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("PedidosPendentes", "RELATORIO") Then
                If ValidaCampos() Then
                    Dim ds As DataSet = getDataSet()
                    ds.Tables(1).TableName = "PedidosPendenteProduto"
                    ds.Tables(2).TableName = "PedidosPendenteCliente"

                    Dim parameters As New Dictionary(Of String, Object)
                    parameters.Add("TipoRelat", ddlClasse.SelectedItem.Text)
                    parameters.Add("ParametrosConsulta", getParameters())
                    parameters.Add("Analitico", chkRelAnalitico.Checked)
                    parameters.Add("Produto", chkRelResumoClienteProduto.Checked)
                    parameters.Add("Cliente", chkRelResumoCliente.Checked)

                    Funcoes.BindReport(Me.Page, ds, "Cr_PedidosPendentes", eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParameters() As String
        Dim param As String = "Empresa" & IIf(ChkConsolidarEmpresa.Checked, " Consolidado: ", ": ") & DdlEmpresa.SelectedItem.Text & vbCrLf
        Dim lst As New List(Of String)

        If chkRelAnalitico.Checked Then
            lst.Add("Analítico")
        End If
        If chkRelResumoClienteProduto.Checked Then
            lst.Add("Cliente/Produto")
        End If
        If chkRelResumoCliente.Checked Then
            lst.Add("Cliente")
        End If
        param &= "Relatórios Selecionados: ('" & String.Join("', '", lst) & "') "

        lst.Clear()
        If chkConsolidarEmpresasApresentacao.Checked OrElse chkConsolidarClientesApresentacao.Checked Then
            If chkConsolidarEmpresasApresentacao.Checked Then
                lst.Add("Consolidar Empresas")
            End If
            If chkConsolidarClientesApresentacao.Checked Then
                lst.Add("Consolidar Clientes")
            End If
            param &= "Apresentação: ('" & String.Join("', '", lst) & "') "
        End If

        If Not String.IsNullOrWhiteSpace(DdlSafra.SelectedValue) Then
            param &= "Safra: " & DdlSafra.SelectedValue & vbCrLf
        End If

        If chkPeriodo.Checked Then
            param &= "Período: " & txtDataInicial.Text & " a " & txtDataFinal.Text
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            param &= " Cliente" & IIf(ChkConsolidarCliente.Checked, " Consolidado: ", ": ") & txtCliente.Text & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            param &= "Produtos: " & ucSelecaoProduto.GetStringGrupoProdutoSelecionado
        End If

        Return param
    End Function
End Class