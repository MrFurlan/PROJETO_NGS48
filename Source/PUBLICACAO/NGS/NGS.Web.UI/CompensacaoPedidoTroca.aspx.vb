Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CompensacaoPedidoTroca
    Inherits BasePage

    Private dateEmissao As New DateTime
    Private intPagina As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CompensacaoPedidoTroca", "ACESSAR") Then
                BuscaEmpresaOrigem()
                BuscaEmpresaDestino()
                CarregarSafrasOrigem()
                CarregarSafrasDestino()
                txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
                txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                txtPedido.Text = ""
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteOrigem" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteOrigem" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteOrigem.Text = itemCliente.Text
            txtCodigoClienteOrigem.Value = itemCliente.Value
            Session.Remove("objClienteOrigem" & HID.Value)
        End If

        If Session("objClienteDestino" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteDestino" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteDestino.Text = itemCliente.Text
            txtCodigoClienteDestino.Value = itemCliente.Value
            Session.Remove("objClienteDestino" & HID.Value)
        End If
    End Sub

    Private Sub BuscaEmpresaOrigem()
        ddl.Carregar(cmbEmpresaOrigem, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub BuscaEmpresaDestino()
        ddl.Carregar(cmbEmpresaDestino, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarSafrasOrigem()
        ddl.Carregar(ddlSafraOrigem, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Sub CarregarSafrasDestino()
        ddl.Carregar(ddlSafraDestino, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If cmbEmpresaOrigem.SelectedIndex = 0 And cmbEmpresaDestino.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione empresa origem ou destino")
            Return False
        ElseIf txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida")
            txtDataInicial.Focus()
            Return False
        ElseIf txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida")
            txtDataFinal.Focus()
            Return False
        ElseIf CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final")
            txtDataInicial.Focus()
            Return False
        End If
        Return True
    End Function

    Protected Sub cmdBuscaClienteOrigem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.SetarHID(HID.Value)
            Popup.ConsultaDeClientes(Me.Page, "objClienteOrigem" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaClienteDestino_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.SetarHID(HID.Value)
            Popup.ConsultaDeClientes(Me.Page, "objClienteDestino" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarRegistros()
        If Funcoes.VerificaPermissao("CompensacaoPedidoTroca", "RELATORIO") Then
            Dim sqlProd As String = ""
            Dim DescricaoProduto As String = ""
            Dim dsProduto As New DataSet
            Dim ParametrosVenda As String = "<-- Venda --> "
            Dim ParametrosCompra As String = "<-- Compra --> "
            Dim PPeriodo As String = ""

            Dim ds As New DataSet
            Dim sql As String
            Dim Cliente As String = ""
            Dim strEmpresaOrigem() As String = cmbEmpresaOrigem.SelectedValue.Split("-")
            Dim strEmpresaDestino() As String = cmbEmpresaDestino.SelectedValue.Split("-")

            Dim strClienteOrigem() As String = txtCodigoClienteOrigem.Value.Split("-")
            Dim strClienteDestino() As String = txtCodigoClienteDestino.Value.Split("-")

            'Moeda exibicao
            If rdOficial.Checked Then
                ParametrosVenda &= "Moeda de Exibição Oficial" & vbCrLf
                ParametrosCompra &= "Moeda de Exibição Oficial" & vbCrLf
            Else
                ParametrosVenda &= "Moeda de Exibição Estrangeira" & vbCrLf
                ParametrosCompra &= "Moeda de Exibição Estrangeira" & vbCrLf
            End If

            sql = "Select distinct" & vbCrLf & _
                  "       P.Empresa_Id                           as EmpresaVenda," & vbCrLf & _
                  "       P.EndEmpresa_id                        as EndEmpresaVenda," & vbCrLf & _
                  "       P.Pedido_Id                            as PedidoVenda," & vbCrLf & _
                  "       P.EmpresaTroca                         as EmpresaCompra," & vbCrLf & _
                  "       P.EndEmpresaTroca                      as EndEmpresaCompra," & vbCrLf & _
                  "       P.PedidoTroca                          as PedidoCompra," & vbCrLf & _
                  "       isnull(PedVenda.ValorOficialVenda,0)   as ValorOficialVenda," & vbCrLf & _
                  "       isnull(PedVenda.ValorMoedaVenda,0)     as ValorMoedaVenda," & vbCrLf & _
                  "       isnull(PedCompra.ValorOficialCompra,0) as ValorOficialCompra," & vbCrLf & _
                  "       isnull(PedCompra.ValorMoedaCompra,0)   as ValorMoedaCompra," & vbCrLf & _
                  "       case" & vbCrLf & _
                  "         when isnull(PedVenda.ValorOficialVenda,0) < isnull(PedCompra.ValorOficialCompra,0)" & vbCrLf & _
                  "           then isnull(PedVenda.ValorOficialVenda,0)" & vbCrLf & _
                  "           else isnull(PedCompra.ValorOficialCompra,0)" & vbCrLf & _
                  "       end CompensarOficial," & vbCrLf & _
                  "       case" & vbCrLf & _
                  "         when isnull(PedVenda.ValorMoedaVenda,0) < isnull(PedCompra.ValorMoedaCompra,0)" & vbCrLf & _
                  "           then isnull(PedVenda.ValorMoedaVenda,0)" & vbCrLf & _
                  "           else isnull(PedCompra.ValorMoedaCompra,0)" & vbCrLf & _
                  "       end CompensarMoeda" & vbCrLf & _
                  "  From Pedidos P" & vbCrLf & _
                  " INNER JOIN PedidoXItem AS PxI" & vbCrLf & _
                  "    ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
                  "	  AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
                  "	  AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
                  " INNER JOIN Produtos AS Pd" & vbCrLf & _
                  "    ON PxI.Produto_Id  = Pd.Produto_Id" & vbCrLf & _
                  " Inner Join Operacoes Op" & vbCrLf & _
                  "    ON Op.Operacao_Id    = P.Operacao" & vbCrLf & _
                  " Inner Join (SELECT Empresa," & vbCrLf & _
                  "                    EndEmpresa," & vbCrLf & _
                  "                    Pedido," & vbCrLf & _
                  "                    sum(isnull(TxC.ValorOficial,0)) as ValorOficialVenda," & vbCrLf & _
                  "                    sum(isnull(TxC.ValorMoeda,0))   as ValorMoedaVenda" & vbCrLf & _
                  "               FROM Titulos AS T" & vbCrLf & _
                  "              Inner Join TitulosxContaContabil TxC" & vbCrLf & _
                  "                 on TxC.Titulo_id = T.Titulo_id" & vbCrLf & _
                  "                and TxC.Conta_Id  = T.ContaContabilRecPag" & vbCrLf & _
                  "              Where T.Situacao = 1 " & vbCrLf & _
                  "                and T.provisao = 3 " & vbCrLf & _
                  "              group by Empresa," & vbCrLf & _
                  "                       EndEmpresa," & vbCrLf & _
                  "                       Pedido" & vbCrLf & _
                  "             ) PedVenda" & vbCrLf & _
                  "    ON P.Empresa_Id    = PedVenda.Empresa" & vbCrLf & _
                  "   and P.EndEmpresa_id = PedVenda.EndEmpresa" & vbCrLf & _
                  "   and P.Pedido_Id     = PedVenda.Pedido" & vbCrLf & _
                  " Inner Join (SELECT Empresa," & vbCrLf & _
                  "                    EndEmpresa," & vbCrLf & _
                  "                    Pedido," & vbCrLf & _
                  "                    sum(isnull(TxC.ValorOficial,0)) as ValorOficialCompra," & vbCrLf & _
                  "                    sum(isnull(TxC.ValorMoeda,0))   as ValorMoedaCompra" & vbCrLf & _
                  "               FROM Titulos AS T" & vbCrLf & _
                  "              Inner Join TitulosxContaContabil TxC" & vbCrLf & _
                  "                 on TxC.Titulo_id = T.Titulo_id" & vbCrLf & _
                  "                and TxC.Conta_Id  = T.ContaContabilRecPag" & vbCrLf & _
                  "              Where T.Situacao = 1" & vbCrLf & _
                  "                and T.provisao = 3" & vbCrLf & _
                  "              group by Empresa," & vbCrLf & _
                  "                       EndEmpresa," & vbCrLf & _
                  "                       Pedido" & vbCrLf & _
                  "             ) PedCompra" & vbCrLf & _
                  "    ON P.EmpresaTroca    = PedCompra.Empresa" & vbCrLf & _
                  "   and P.EndEmpresaTroca = PedCompra.EndEmpresa" & vbCrLf & _
                  "   and P.PedidoTroca     = PedCompra.Pedido" & vbCrLf & _
                  " Where P.Situacao        =  1" & vbCrLf & _
                  "   And Op.Classe         = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
                  "   And isnull(P.Troca,0) = 1" & vbCrLf & _
                  "   And ValorOficialVenda  > 0" & vbCrLf & _
                  "   And ValorOficialCompra > 0" & vbCrLf

            If txtPedido.Text.Trim <> "" Then
                ParametrosVenda &= "Pedido: " & txtPedido.Text & vbCrLf
                sql &= " AND P.Pedido_Id   = " & txtPedido.Text & vbCrLf
            ElseIf txtPedidoCompra.Text.Trim <> "" Then
                ParametrosCompra &= "Pedido: " & txtPedido.Text & vbCrLf
                sql &= " AND P.PedidoTroca =  " & txtPedido.Text & vbCrLf
            Else
                If ChkPeriodo.Checked Then
                    PPeriodo = "Periodo de: " & txtDataInicial.Text.ToStrDate() & " a " & txtDataFinal.Text.ToString()
                    sql &= " and P.DataEntrega BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate & "' " & vbCrLf
                End If

                If strEmpresaOrigem(0).Length > 0 Then
                    ParametrosVenda &= "Empresa: " & cmbEmpresaOrigem.SelectedItem.Text & vbCrLf
                    sql &= " AND P.Empresa_Id    = '" & strEmpresaOrigem(0) & "'" & vbCrLf
                    sql &= " AND P.EndEmpresa_Id = " & strEmpresaOrigem(1) & " " & vbCrLf
                End If

                If strClienteOrigem(0).Length > 0 Then
                    If chkConsClienteVenda.Checked Then
                        ParametrosVenda &= "Cliente Consolidado"
                        sql &= " AND left(P.Cliente,8) = '" & strClienteOrigem(0).ToString.Substring(0, 8) & "'" & vbCrLf
                    Else
                        sql &= " AND P.Cliente= '" & strClienteOrigem(0) & "'" & vbCrLf
                        sql &= " AND P.EndCliente= " & strClienteOrigem(1) & " " & vbCrLf
                    End If
                    ParametrosVenda &= "Cliente: " & txtClienteOrigem.Text & vbCrLf
                End If

                If ucSelecaoProdutoVenda.TemSelecionado Then
                    Dim ret As New ArrayList
                    ret = ucSelecaoProdutoVenda.GetSqlEParametrosRelatorio("Pd.Grupo", "PxI.Produto_Id", "", True)
                    sql &= " and " & ret(0) & vbCrLf
                    ParametrosVenda &= ret(1)
                End If

                If ddlSafraOrigem.SelectedIndex > 0 Then
                    ParametrosVenda &= "Safra: " & ddlSafraOrigem.SelectedItem.Text & vbCrLf
                    sql &= " AND P.Safra = '" & ddlSafraOrigem.SelectedValue & "'" & vbCrLf
                End If

                If strEmpresaDestino(0).Length > 0 Or strClienteDestino(0).Length > 0 Or ucSelecaoProdutoCompra.TemSelecionado Or ddlSafraDestino.SelectedIndex > 0 Then
                    sql &= " And P.Pedido_Id in (Select P.PedidoTroca" & vbCrLf & _
                           "					   From Pedidos P" & vbCrLf & _
                           "				      INNER JOIN PedidoXItem AS PxI" & vbCrLf & _
                           " 				         ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
                           "				  	    AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
                           "				  	    AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
                           "				      INNER JOIN Produtos AS Pd" & vbCrLf & _
                           "				  	     ON PxI.Produto_Id  = Pd.Produto_Id" & vbCrLf & _
                           "                      Inner Join SubOperacoes SO" & vbCrLf & _
                           "                         ON SO.Operacao_Id     = P.Operacao" & vbCrLf & _
                           "                        And SO.SubOperacoes_Id = P.SubOperacao" & vbCrLf & _
                           "				      Where P.Situacao        = 1 " & vbCrLf & _
                           "                        And isnull(P.troca,0) = 1 "


                    If strEmpresaDestino(0).Length > 0 Then
                        ParametrosCompra &= "Empresa: " & cmbEmpresaDestino.SelectedItem.Text & vbCrLf
                        sql &= "                        AND P.Empresa_Id    ='" & strEmpresaDestino(0) & "'" & vbCrLf
                        sql &= "                        AND P.EndEmpresa_Id = " & strEmpresaDestino(1) & " " & vbCrLf
                    End If

                    If strClienteDestino(0).Length > 0 Then
                        If chkConsClienteCompra.Checked Then
                            ParametrosCompra &= "Cliente Consolidado"
                            sql &= "                        AND left(P.Cliente,8)    ='" & strClienteDestino(0).ToString.Substring(0, 8) & "'" & vbCrLf
                        Else
                            sql &= "                        AND P.Cliente    ='" & strClienteDestino(0) & "'" & vbCrLf
                            sql &= "                        AND P.EndCliente = " & strClienteDestino(1) & " " & vbCrLf
                        End If
                        ParametrosCompra &= "Cliente: " & txtClienteDestino.Text & vbCrLf
                    End If

                    If ucSelecaoProdutoCompra.TemSelecionado Then
                        Dim ret As New ArrayList
                        ret = ucSelecaoProdutoCompra.GetSqlEParametrosRelatorio("Pd.Grupo", "PxI.Produto_Id", "", True)
                        sql &= "                        And " & ret(0) & vbCrLf
                        ParametrosCompra &= ret(1)
                    End If

                    If ddlSafraDestino.SelectedIndex > 0 Then
                        ParametrosCompra &= "Safra: " & ddlSafraDestino.SelectedItem.Text & vbCrLf
                        sql &= "                        AND P.Safra = '" & ddlSafraDestino.SelectedValue & "'" & vbCrLf
                    End If
                    sql &= "					)"
                End If
                ParametrosCompra &= PPeriodo
            End If

            ds = Banco.ConsultaDataSet(sql, "Consulta")
            gdvPedido.DataSource = ds
            gdvPedido.DataBind()
            If gdvPedido.Rows.Count > 0 Then TabContainer1.ActiveTabIndex = 2
        Else
            MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
        End If
    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal Caminho As String, ByVal ParametrosV As String, ByVal ParametrosC As String)
        Try
            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("ParametrosV", ParametrosV)
            parameters.Add("ParametrosC", ParametrosC)
            parameters.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
            parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))
            parameters.Add("Zerado", "True")
            parameters.Add("Titulo", "Relatório de Vendas Por Troca")
            parameters.Add("Pagina", "Página")
            parameters.Add("Data", "Emissão")

            parameters.Add("Pedido", "Pedido")
            parameters.Add("Produto", "Produto")
            parameters.Add("Cliente", "Cliente")
            parameters.Add("Moeda", "Moeda")
            parameters.Add("Quantidade", "Quantidade")
            parameters.Add("ValorOficial", IIf(rdOficial.Checked, "Vlr.Oficial", "Vlr.Moeda"))
            parameters.Add("TotalOficial", IIf(rdOficial.Checked, "Total Ofic.", "Total Moeda"))

            parameters.Add("EntregaNota", "Entr.Nota")
            parameters.Add("DevolucaoNota", "Dev. Nota")
            parameters.Add("EntregaValorNota", "Ent. Vlr Nota")
            parameters.Add("DevolucaoValorNota", "Dev. Vlr Nota")

            Funcoes.BindReport(Me.Page, Ds, Caminho, eExportType.PDF, parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If cmbEmpresaOrigem.SelectedIndex = -1 Or txtCodigoClienteOrigem.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e/ou cliente para buscar o número do pedido")
            Else
                Dim strJScript As String = ""
                Dim strCodigoEmpresaOrigem As String() = cmbEmpresaOrigem.SelectedValue.Split("-")
                Dim strCodClienteOrigem As String() = txtCodigoClienteOrigem.Value.Split(";")

                HttpContext.Current.Session("ssCampo") = "ApenasNumeroPedido"
                HttpContext.Current.Session("ssCnpjDaEmpresa") = strCodigoEmpresaOrigem(0)
                HttpContext.Current.Session("ssEndDaEmpresa") = strCodigoEmpresaOrigem(1)
                HttpContext.Current.Session("txtCnpjDoCliente") = strCodClienteOrigem(0)
                HttpContext.Current.Session("txtEndDoCliente") = strCodClienteOrigem(1)

                'HttpContext.Current.Session("CodigoProduto") = cmbProdutoOrigem.SelectedValue
                HttpContext.Current.Session("CodigoSafra") = ddlSafraOrigem.SelectedValue

                strJScript += "var x = (screen.height / 2) - 250; "
                strJScript += "var y = (screen.width / 2) - 400; "
                strJScript += "window.open(""ConsultaPedidos.aspx?pedido=" & txtPedido.ClientID & "&tipo=v"", """", ""resizable=no, menubar=no, scrollbars=yes, width=800, height=500, top="" + x + "", left="" + y + """");"
                ScriptManager.RegisterClientScriptBlock(Me, cmdBuscaPedido.GetType(), "Extratos", strJScript, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedidoCompra_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If cmbEmpresaDestino.SelectedIndex = -1 Or txtCodigoClienteDestino.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e/ou cliente para buscar o número do pedido")
            Else
                Dim strJScript As String = ""
                Dim strCodigoEmpresaDestino As String() = cmbEmpresaDestino.SelectedValue.Split("-")
                Dim strCodClienteDestino As String() = txtCodigoClienteDestino.Value.Split(";")

                HttpContext.Current.Session("ssCampo") = "ApenasNumeroPedido"
                HttpContext.Current.Session("ssCnpjDaEmpresa") = strCodigoEmpresaDestino(0)
                HttpContext.Current.Session("ssEndDaEmpresa") = strCodigoEmpresaDestino(1)
                HttpContext.Current.Session("txtCnpjDoCliente") = strCodClienteDestino(0)
                HttpContext.Current.Session("txtEndDoCliente") = strCodClienteDestino(1)

                'HttpContext.Current.Session("CodigoProduto") = cmbProdutoDestino.SelectedValue
                HttpContext.Current.Session("CodigoSafra") = ddlSafraDestino.SelectedValue

                strJScript += "var x = (screen.height / 2) - 250; "
                strJScript += "var y = (screen.width / 2) - 400; "
                strJScript += "window.open(""ConsultaPedidos.aspx?pedido=" & txtPedidoCompra.ClientID & "&tipo=v"", """", ""resizable=no, menubar=no, scrollbars=yes, width=800, height=500, top="" + x + "", left="" + y + """");"
                ScriptManager.RegisterClientScriptBlock(Me, cmdBuscaPedido.GetType(), "Extratos", strJScript, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ChkPeriodo.CheckedChanged
        Try
            pnlData.Visible = ChkPeriodo.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            txtClienteOrigem.Text = ""
            txtCodigoClienteOrigem.Value = ""
            txtClienteDestino.Text = ""
            txtCodigoClienteDestino.Value = ""
            ddlSafraOrigem.SelectedIndex = 0
            ddlSafraDestino.SelectedIndex = 0
            txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtPedido.Text = ""
            cmbEmpresaOrigem.SelectedIndex = 0
            cmbEmpresaDestino.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ValidarCampos() Then
                BuscarRegistros()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkCompensar_Click(sender As Object, e As EventArgs) Handles LnkCompensar.Click
        Try
            Compensar()
            BuscarRegistros()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString)
        End Try
    End Sub

    Protected Sub Compensar()
        Dim CodigoPedido As Integer = 0
        Dim Empresa As String = String.Empty
        Dim Endereco As String = String.Empty
        Dim Sqls As New ArrayList
        Dim vlrParaCompensar As Decimal = Decimal.Zero
        Dim Num As New [Lib].Negocio.Numerador(1)
        Dim i As Integer = 0

        For Each row As GridViewRow In gdvPedido.Rows
            If CType(row.FindControl("chkPedido"), CheckBox).Checked Then
                'Compensa sempre o de menor valor em provisão
                If CDec(row.Cells(7).Text) > CDec(row.Cells(9).Text) Then
                    CodigoPedido = row.Cells(3).Text 'Venda
                    Empresa = row.Cells(1).Text
                    Endereco = row.Cells(2).Text
                Else
                    CodigoPedido = row.Cells(6).Text 'Compra
                    Empresa = row.Cells(4).Text
                    Endereco = row.Cells(5).Text
                End If
                'Recupera o Pedido
                Dim Pedido As New Pedido(Empresa, Endereco, CodigoPedido)

                For j As Integer = 0 To 100 Step 1
                    Dim tit As Novo.TituloNovo = Pedido.Titulos(j)
                    If tit.CodigoProvisao = eProvisao.Provisao Then
                        'Pedido
                        vlrParaCompensar = tit.Valores.EncargoValorDocumento.Valor
                        If String.IsNullOrEmpty(tit.IUD) Then tit.IUD = "U"
                        tit.CodigoContaContabilRecPag = Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao
                        'Pedido de troca
                        For k As Integer = 0 To 100 Step 1
                            Dim titTroca As Novo.TituloNovo = Pedido.PedidoTroca.Titulos(0)
                            If titTroca.CodigoProvisao = eProvisao.Provisao AndAlso titTroca.CodigoTituloCompensacao = 0 AndAlso vlrParaCompensar > 0 Then
                                tit.CodigoProvisao = eProvisao.Baixa 'baixa o titulo que vai compensar.
                                tit.Sequencia = tit.Codigo
                                If String.IsNullOrEmpty(titTroca.IUD) Then titTroca.IUD = "U"
                                titTroca.CodigoContaContabilRecPag = Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao
                                titTroca.CodigoProvisao = eProvisao.Baixa
                                titTroca.Sequencia = titTroca.Codigo
                                titTroca.CodigoTituloCompensacao = tit.Codigo 'Título de Compensação

                                If vlrParaCompensar > titTroca.Valores.EncargoValorDocumento.Valor Then
                                    vlrParaCompensar -= titTroca.Valores.EncargoValorDocumento.Valor
                                ElseIf vlrParaCompensar < titTroca.Valores.EncargoValorDocumento.Valor Then
                                    Dim TituloTroca As New Novo.TituloNovo(titTroca.Codigo)
                                    TituloTroca.IUD = "I"
                                    i += 1
                                    TituloTroca.Codigo = Num.Sequencia + i
                                    TituloTroca.Valores.EncargoValorDocumento.Valor = titTroca.Valores.EncargoValorDocumento.Valor - vlrParaCompensar
                                    TituloTroca.CodigoProvisao = eProvisao.Provisao
                                    TituloTroca.NotaTitulo.IUD = "I"
                                    Pedido.PedidoTroca.Titulos.Add(TituloTroca)
                                    titTroca.Valores.EncargoValorDocumento.Valor = vlrParaCompensar
                                    vlrParaCompensar = 0
                                End If
                                titTroca.SalvarSql(Sqls, False, True)
                            End If
                            If Pedido.PedidoTroca.Titulos.Count - 1 = k Then Exit For
                        Next

                        If vlrParaCompensar > 0 AndAlso tit.Valores.EncargoValorDocumento.Valor > vlrParaCompensar Then
                            Dim Titulo As New Novo.TituloNovo(tit.Codigo)
                            Titulo.IUD = "I"
                            i += 1
                            Titulo.Codigo = Num.Sequencia + i
                            Titulo.Valores.EncargoValorDocumento.Valor = vlrParaCompensar
                            Titulo.CodigoProvisao = eProvisao.Provisao
                            Titulo.NotaTitulo.IUD = "I"
                            Pedido.Titulos.Add(Titulo)
                            tit.Valores.EncargoValorDocumento.Valor -= vlrParaCompensar
                        End If

                        tit.SalvarSql(Sqls, False, True)
                        vlrParaCompensar = 0
                    End If
                    If Pedido.Titulos.Count - 1 = j Then Exit For
                Next
                If Pedido.Titulos.Count > 0 Then Exit For
            End If
        Next

        If i > 0 Then Sqls.Add(Num.IncrementarNumeradorSql(True, i + 1))

        If Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, "Compensação realizada com Sucesso.", eTitulo.Sucess)
        Else
            MsgBox(Me.Page, "Não foi possível realizar a compensação")
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CompensacaoPedidoTroca")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class