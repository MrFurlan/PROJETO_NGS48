Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class PedidosEntreguesDeVendas
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PedidosEntreguesDeVendas", "ACESSAR") Then
                    CarregarTipoDoItem()
                    CargaUnidade()
                    Limpar()
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    ddl.Carregar(ddlGruposCFOP, CarregarDDL.Tabela.CFOPGrupo, "", True)
                    ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                    ddl.Carregar(ddlEstadoFisico, CarregarDDL.Tabela.EstadoFisicoIA, "", True)
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

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
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

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Validar() Then
                Dim ds As DataSet = getDataSet(False)

                mostrargrid()

                If RadCliente.Checked = True Then
                    GridCliente.DataSource = ds
                    GridCliente.DataBind()
                Else
                    GridProduto.DataSource = ds
                    GridProduto.DataBind()
                End If
            End If
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

#End Region

#Region "Methods"

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, " Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario"))
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub CarregarTipoDoItem()
        lstTipoDoItem.Items.Add(New ListItem("00 - Mercadorias para Revenda", 0))
        lstTipoDoItem.Items.Add(New ListItem("01 - Matéria Prima", 1))
        lstTipoDoItem.Items.Add(New ListItem("02 - Embalagem", 2))
        lstTipoDoItem.Items.Add(New ListItem("03 - Produto em Processo", 3))
        lstTipoDoItem.Items.Add(New ListItem("04 - Produto Acabado", 4))
        lstTipoDoItem.Items.Add(New ListItem("05 - SubProduto", 5))
        lstTipoDoItem.Items.Add(New ListItem("06 - Produto Intermediário", 6))
        lstTipoDoItem.Items.Add(New ListItem("07 - Material de Uso e Consumo", 7))
        lstTipoDoItem.Items.Add(New ListItem("08 - Ativo Imobilizado", 8))
        lstTipoDoItem.Items.Add(New ListItem("09 - Serviços", 9))
        lstTipoDoItem.Items.Add(New ListItem("10 - Outros Insumos", 10))
        lstTipoDoItem.Items.Add(New ListItem("99 - Outras", 99))
    End Sub

    Protected Sub ddlGruposCFOP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlGruposCFOP.SelectedIndex > 0 Then
                Panel3.Visible = True
                Dim ListaCFOP As New [Lib].Negocio.ListCFOP(True, ddlGruposCFOP.SelectedValue)

                lstCfop.Items.Clear()
                lstCfopSelecionados.Items.Clear()

                Dim j As Integer = 0
                While j < ListaCFOP.Count
                    lstCfop.Items.Add(New ListItem(Format(ListaCFOP(j).Codigo, "0000") & "-" & ListaCFOP(j).Descricao, ListaCFOP(j).Codigo))
                    j += 1
                End While
            Else
                Panel3.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgAdicionar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If lstCfop.SelectedIndex > -1 Then

                Dim temCFOP As Boolean = False

                Dim i As Integer = 0
                While i < lstCfopSelecionados.Items.Count
                    If lstCfopSelecionados.Items(i).Text = lstCfop.SelectedItem.Text Then
                        temCFOP = True
                    End If
                    i += 1
                End While

                If temCFOP Then
                    MsgBox(Me.Page, "CFOP já foi selecionado")
                Else
                    lstCfopSelecionados.Items.Add(New ListItem(lstCfop.SelectedItem.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemover_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If lstCfopSelecionados.SelectedIndex > -1 Then
                lstCfopSelecionados.Items.RemoveAt(lstCfopSelecionados.SelectedIndex)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        VerificaUnidade()
        txtDataInicial.Text = New DateTime(Year(DateTime.Now), Month(DateTime.Now), 1).ToString("dd/MM/yyyy")
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
        lstTipoDoItem.SelectedIndex = -1
        RadCliente.Checked = True
        ckApenasFinanceiro.Checked = True
        ucSelecaoProduto.Limpar()
        ddlGruposCFOP.SelectedIndex = 0
        Panel3.Visible = False
        lstCfop.Items.Clear()
        lstCfopSelecionados.Items.Clear()
        mostrargrid()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub mostrargrid()
        GridCliente.DataBind()
        GridProduto.DataBind()

        If RadCliente.Checked Then
            GridCliente.Parent.Visible = True
            GridProduto.Parent.Visible = False
        Else
            GridProduto.Parent.Visible = True
            GridCliente.Parent.Visible = False
        End If
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.Text) Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe uma data inicial válida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataFinal.Text) OrElse Not IsDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe uma data final válida.")
            Return False
        ElseIf CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final.")
            Return False
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

    Private Function getListTipoDoItem(ByVal pos As Integer) As List(Of String)
        Try
            Dim lst As New List(Of String)

            For Each item As String In lstTipoDoItem.GetSelectedValues()
                lst.Add(item.Split("-")(pos))
            Next
            Return lst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSet(ByVal eDados As Boolean) As DataSet
        Dim Sql As String = String.Empty

        'Por Cliente
        If RadCliente.Checked = True Then
            Sql = " SELECT NotasFiscais.Empresa_Id as Empresa, " & vbCrLf &
                  " NotasFiscais.Nota_Id as Nota,  " & vbCrLf &
                  "        NotasFiscais.Movimento as Data,  " & vbCrLf &
                  "        Clientes.Nome AS Cliente, " & vbCrLf &
                  "        ISNULL(rep.Nome,'') AS Representante, " & vbCrLf &
                  "        Clientes.Bairro,  " & vbCrLf &
                  "        Clientes.Cidade AS Cidade,      " & vbCrLf &
                  "        Clientes.Estado AS Estado,  " & vbCrLf &
                  "        sum(Case " & vbCrLf &
                  "			when (Notasfiscais.Operacao=70 and Notasfiscais.suboperacao=1) " & vbCrLf &
                  "				then 0 " & vbCrLf &
                  "				else NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf &
                  "			end) AS Quantidade,  " & vbCrLf &
                  "        sum(NFPrd.ValorPrd) as ValorPrd, " & vbCrLf &
                  "        sum(NFEnc.ValorNF) as Valor " & vbCrLf &
                  "   FROM NotasFiscais  " & vbCrLf &
                  "  INNER JOIN NotasFiscaisXItens  " & vbCrLf &
                  "     ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id  " & vbCrLf &
                  "    AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf &
                  "    AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id  " & vbCrLf &
                  "    AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf &
                  "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf &
                  "    AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                  "    AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id  " & vbCrLf &
                  "		  INNER JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "							EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, " & vbCrLf &
                  "							sum(Valor) as ValorPrd " & vbCrLf &
                  "					FROM NotasFiscaisXEncargos " & vbCrLf &
                  "					WHERE Encargo_ID IN('PRODUTO','FRETES','DESP.ADUANEIRAS','SEGURO') " & vbCrLf &
                  "					GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "								EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id) as NFPrd " & vbCrLf &
                  "				 ON NotasFiscaisXItens.Empresa_Id      = NFPrd.Empresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndEmpresa_Id   = NFPrd.EndEmpresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Cliente_Id      = NFPrd.Cliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndCliente_Id   = NFPrd.EndCliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EntradaSaida_Id = NFPrd.EntradaSaida_Id  " & vbCrLf &
                  "				AND NotasFiscaisXItens.Serie_Id        = NFPrd.Serie_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Nota_Id         = NFPrd.Nota_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Produto_Id      = NFPrd.Produto_Id " & vbCrLf &
                  "		  INNER JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "							EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, " & vbCrLf &
                  "							sum(case  " & vbCrLf &
                  "									when Encargo_ID ='DESCONTOS' " & vbCrLf &
                  "										then Valor * -1 " & vbCrLf &
                  "										else Valor " & vbCrLf &
                  "									end) as ValorNF " & vbCrLf &
                  "					FROM NotasFiscaisXEncargos " & vbCrLf &
                  "					WHERE Encargo_ID IN('PRODUTO','FRETES','DESP.ADUANEIRAS','DESCONTOS','SEGURO') " & vbCrLf &
                  "					GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "								EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id) as NFEnc " & vbCrLf &
                  "				 ON NotasFiscaisXItens.Empresa_Id      = NFEnc.Empresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndEmpresa_Id   = NFEnc.EndEmpresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Cliente_Id      = NFEnc.Cliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndCliente_Id   = NFEnc.EndCliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EntradaSaida_Id = NFEnc.EntradaSaida_Id  " & vbCrLf &
                  "				AND NotasFiscaisXItens.Serie_Id        = NFEnc.Serie_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Nota_Id         = NFEnc.Nota_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Produto_Id      = NFEnc.Produto_Id " & vbCrLf &
                  "  INNER JOIN SubOperacoes SO " & vbCrLf &
                  "     ON SO.Operacao_Id      = Notasfiscais.Operacao " & vbCrLf &
                  "    AND SO.Suboperacoes_Id = notasfiscais.suboperacao " & vbCrLf &
                  "  INNER JOIN Clientes  " & vbCrLf &
                  "     ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id " & vbCrLf &
                  "    AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id  " & vbCrLf &
                  "  INNER JOIN Pedidos ON NotasFiscais.Empresa_Id = Pedidos.Empresa_Id  " & vbCrLf &
                  "    AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id  " & vbCrLf &
                  "    AND NotasFiscais.Pedido = Pedidos.Pedido_Id  " & vbCrLf &
                  "  INNER JOIN Produtos P " & vbCrLf &
                  "     ON NotasFiscaisXItens.Produto_Id = P.Produto_Id     " & vbCrLf
            Sql &= " LEFT JOIN Comissoes CO ON                              " & vbCrLf &
                   "        CO.Empresa_Id = Pedidos.Empresa_Id              " & vbCrLf &
                   "        AND CO.EndEmpresa_Id = Pedidos.EndEmpresa_Id    " & vbCrLf &
                   "        And CO.Pedido_Id = pedidos.Pedido_Id            " & vbCrLf &
                   " LEFT JOIN clientes rep ON                              " & vbCrLf &
                   "        rep.Cliente_Id = co.Representante_Id            " & vbCrLf &
                   "        AND rep.Endereco_Id = co.EndRepresentante_Id    " & vbCrLf

            Sql &= whereSQL()

            Sql &= " GROUP  BY NotasFiscais.Empresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
                   "        NotasFiscais.Movimento, Clientes.Nome, ISNULL(rep.Nome,''), Clientes.Bairro, Clientes.Cidade, Clientes.Estado, Pedidos.PedidoEfetivo" & vbCrLf &
                   " ORDER  BY Empresa, Cliente, NotasFiscais.Nota_Id" & vbCrLf

        Else 'Por Produto

            Sql = " SELECT NotasFiscais.Empresa_Id as Empresa, " & vbCrLf &
                  " NotasFiscais.Nota_Id as Nota,  " & vbCrLf &
                  "        NotasFiscais.Movimento AS Data,  " & vbCrLf &
                  "        Clientes.Nome AS Cliente,  " & vbCrLf &
                  "        ISNULL(rep.Nome, '') AS Representante,  " & vbCrLf &
                  "        P.NCM," & vbCrLf &
                  "        NotasFiscaisXItens.Produto_Id as Produto,  " & vbCrLf &
                  "        P.Unidade, " & vbCrLf &
                  "        P.Nome as NomeDoProduto, " & vbCrLf &
                  "        ISNULL(NxL.Lote_Id,'') Lote, " & vbCrLf &
                  "        ISNULL(CONVERT(VARCHAR,NxL.Validade,103),'') AS Validade, " & vbCrLf

            If eDados Then
                Sql &= "        EstadoFisicoIA.Descricao AS EstadoFisico,"
            End If

            Sql &= "        Case " & vbCrLf &
                  "			when (Notasfiscais.Operacao=70 and Notasfiscais.suboperacao=1) " & vbCrLf &
                  "				then 0 " & vbCrLf &
                  "				else NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf &
                  "			end AS Quantidade,  " & vbCrLf &
                  "        NotasFiscaisXItens.Unitario, NFEnc.ValorNF as Valor, " & vbCrLf &
                  "        NFPrd.ValorPrd as ValorPrd, " & vbCrLf &
                  "        isnull(NotasFiscaisXEncargos.Percentual,0) as PercentualIcms, " & vbCrLf &
                  "        isnull(NotasFiscaisXEncargos.Valor,0) as ValorIcms " & vbCrLf &
                  "   FROM NotasFiscais  " & vbCrLf &
                  "		  INNER JOIN NotasFiscaisXItens  " & vbCrLf &
                  "				 ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id   " & vbCrLf &
                  "				AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                  "				AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf &
                  "				AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                  "				AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf &
                  "				AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                  "				AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                  "		  LEFT JOIN NotasFiscaisXEncargos  " & vbCrLf &
                  "				 ON NotasFiscaisXEncargos.Empresa_Id      = NotasFiscaisXItens.Empresa_Id   " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.Cliente_Id      = NotasFiscaisXItens.Cliente_Id   " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.Produto_Id      = NotasFiscaisXItens.Produto_Id " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.CFOP_Id         = NotasFiscaisXItens.CFOP_Id " & vbCrLf &
                  "				AND NotasFiscaisXEncargos.Encargo_id      = 'ICMS' " & vbCrLf &
                  "		  INNER JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "							EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, " & vbCrLf &
                  "							sum(Valor) as ValorPrd " & vbCrLf &
                  "					FROM NotasFiscaisXEncargos " & vbCrLf &
                  "					WHERE Encargo_ID IN('PRODUTO','FRETES','DESP.ADUANEIRAS','SEGURO') " & vbCrLf &
                  "					GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "								EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id) as NFPrd " & vbCrLf &
                  "				 ON NotasFiscaisXItens.Empresa_Id      = NFPrd.Empresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndEmpresa_Id   = NFPrd.EndEmpresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Cliente_Id      = NFPrd.Cliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndCliente_Id   = NFPrd.EndCliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EntradaSaida_Id = NFPrd.EntradaSaida_Id  " & vbCrLf &
                  "				AND NotasFiscaisXItens.Serie_Id        = NFPrd.Serie_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Nota_Id         = NFPrd.Nota_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Produto_Id      = NFPrd.Produto_Id " & vbCrLf &
                  "		  INNER JOIN (SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "							EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, " & vbCrLf &
                  "							sum(case  " & vbCrLf &
                  "									when Encargo_ID ='DESCONTOS' " & vbCrLf &
                  "										then Valor * -1 " & vbCrLf &
                  "										else Valor " & vbCrLf &
                  "									end) as ValorNF " & vbCrLf &
                  "					FROM NotasFiscaisXEncargos " & vbCrLf &
                  "					WHERE Encargo_ID IN('PRODUTO','FRETES','DESP.ADUANEIRAS','DESCONTOS', 'SEGURO') " & vbCrLf &
                  "					GROUP BY Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf &
                  "								EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id) as NFEnc " & vbCrLf &
                  "				 ON NotasFiscaisXItens.Empresa_Id      = NFEnc.Empresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndEmpresa_Id   = NFEnc.EndEmpresa_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Cliente_Id      = NFEnc.Cliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EndCliente_Id   = NFEnc.EndCliente_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.EntradaSaida_Id = NFEnc.EntradaSaida_Id  " & vbCrLf &
                  "				AND NotasFiscaisXItens.Serie_Id        = NFEnc.Serie_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Nota_Id         = NFEnc.Nota_Id " & vbCrLf &
                  "				AND NotasFiscaisXItens.Produto_Id      = NFEnc.Produto_Id " & vbCrLf &
                  "		  INNER JOIN SubOperacoes SO " & vbCrLf &
                  "				 ON SO.Operacao_Id      = Notasfiscais.Operacao " & vbCrLf &
                  "				AND SO.Suboperacoes_Id = notasfiscais.suboperacao " & vbCrLf &
                  "		  INNER JOIN Clientes " & vbCrLf &
                  "				 ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id " & vbCrLf &
                  "				AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf &
                  "		  INNER JOIN Pedidos ON NotasFiscais.Empresa_Id = Pedidos.Empresa_Id  " & vbCrLf &
                  "				AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id  " & vbCrLf &
                  "				AND NotasFiscais.Pedido = Pedidos.Pedido_Id  " & vbCrLf &
                  "		  INNER JOIN Produtos P ON NotasFiscaisXItens.Produto_Id = P.Produto_Id " & vbCrLf &
                  "		  LEFT JOIN EstadoFisicoIA ON EstadoFisicoIA.EstadoFisicoIA_Id = P.CodigoEstadoFisico " & vbCrLf &
                  "       LEFT JOIN NotaFiscalXLote NxL On NxL.Empresa_Id = NotasFiscaisXItens.Empresa_Id " & vbCrLf &
                  "             AND NxL.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                  "             AND NxL.Cliente_Id = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
                  "             AND NxL.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                  "             AND NxL.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                  "             AND NxL.Serie_Id = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                  "             AND NxL.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
                  "             AND NxL.Produto_Id = NotasFiscaisXItens.Produto_Id " & vbCrLf &
                  "             AND NxL.Sequencia_Id = NotasFiscaisXItens.Sequencia_Id " & vbCrLf &
                  "             AND NxL.CFOP_Id = NotasFiscaisXItens.CFOP_Id        " & vbCrLf &
                   "        LEFT JOIN Comissoes CO ON                               " & vbCrLf &
                   "               CO.Empresa_Id = Pedidos.Empresa_Id               " & vbCrLf &
                   "               AND CO.EndEmpresa_Id = Pedidos.EndEmpresa_Id     " & vbCrLf &
                   "               And CO.Pedido_Id = pedidos.Pedido_Id             " & vbCrLf &
                   "        LEFT JOIN clientes rep ON                               " & vbCrLf &
                   "               rep.Cliente_Id = co.Representante_Id             " & vbCrLf &
                   "               AND rep.Endereco_Id = co.EndRepresentante_Id     " & vbCrLf

            Sql &= whereSQL()

            Sql &= " ORDER  BY Empresa, Cliente, Nota " & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
        Return ds
    End Function

    Private Function whereSQL() As String
        whereSQL = String.Empty
        Dim par As ArrayList

        whereSQL &= " WHERE  NotasFiscais.tipodedocumento = 1 And (NotasFiscais.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "')  " & vbCrLf & _
                  " AND (NotasFiscais.EntradaSaida_Id = 'S')  " & vbCrLf & _
                  " AND (NotasFiscais.Situacao = 1)  " & vbCrLf & _
                  " AND (NotasFiscais.NFG = 0)  " & vbCrLf


        If Not chkEmpresa.Checked AndAlso DdlEmpresa.SelectedIndex > 0 Then
            Dim strEmpresa As String()
            strEmpresa = DdlEmpresa.SelectedValue.Split(New Char() {"-"})
            whereSQL &= "AND  NotasFiscais.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf &
                         "AND  NotasFiscais.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
        End If

        Dim lst As List(Of String) = getListTipoDoItem(0)

        If lst.Count > 0 Then
            whereSQL &= " And P.TipoDoItem IN ( " & String.Join(", ", lst) & ")"
        End If

        If ckApenasFinanceiro.Checked Then
            whereSQL &= " AND SO.Financeiro = 'S' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            whereSQL &= "  And (NotasFiscais.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "') And NotasFiscais.EndCliente_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        End If

        If cmbOperacao.SelectedIndex > 0 Then whereSQL &= "AND NotasFiscais.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf

        If cmbSubOperacao.SelectedIndex > 0 Then
            Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            whereSQL &= "AND NotasFiscais.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
        End If

        If RadProduto.Checked AndAlso ddlEstadoFisico.SelectedIndex > 0 Then
            whereSQL &= "AND P.CodigoEstadoFisico = " & ddlEstadoFisico.SelectedValue & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado() Then
            par = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "P.Produto_Id")
            whereSQL &= " AND " & par(0)
        End If

        If lstCfopSelecionados.Items.Count > 0 Then
            whereSQL &= "AND ("
            Dim strOr As String = ""

            Dim k As Integer = 0
            While k < lstCfopSelecionados.Items.Count
                whereSQL &= strOr & "NotasFiscaisXItens.CFOP_Id = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                strOr = " OR "
                k += 1
            End While

            whereSQL &= ") " & vbCrLf
        End If

    End Function

#End Region

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

    Protected Sub lnkExcelD_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelD.Click
        Try
            If RadProduto.Checked Then
                EmitirRelatorioDados()
            Else
                MsgBox(Me.Page, "Apenas para Relatório por Produto.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("PedidosEntreguesDeVendas", "RELATORIO") Then
                If Validar() Then
                    Dim ds As DataSet = getDataSet(False)

                    Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Nome", objEmpresa.Nome)
                    parameters.Add("Cidade", String.Format("{0}-{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado))
                    parameters.Add("Zerado", "True")
                    parameters.Add("Adiantamento", "")
                    parameters.Add("Saldo", "")

                    parameters.Add("Titulo", "Relatório de Pedidos Entregues De Vendas ")
                    parameters("Titulo") &= IIf(RadCliente.Checked, "(Por Cliente)", "(Por Produto)") & " - De: " & txtDataInicial.Text & " Até: " & txtDataFinal.Text
                    parameters.Add("Numero", IIf(RadCliente.Checked, "", "NOTA"))
                    parameters.Add("Romaneio", IIf(RadCliente.Checked, "QUANTIDADE", ""))
                    parameters.Add("Produto", IIf(RadCliente.Checked, "CLIENTE", "PRODUTO"))
                    parameters.Add("Origem", IIf(RadCliente.Checked, "BAIRRO", ""))
                    parameters.Add("Destino", IIf(RadCliente.Checked, "DATA", "CLIENTE"))
                    parameters.Add("Placa", IIf(RadCliente.Checked, "NOTA", "DATA"))
                    parameters.Add("Transportador", IIf(RadCliente.Checked, "CIDADE", "UNIDADE"))
                    parameters.Add("Motorista", IIf(RadCliente.Checked, "UF", ""))
                    parameters.Add("Quantidade", IIf(RadCliente.Checked, "", "QUANTIDADE"))
                    'parameters.Add("ValorPrd", IIf(RadCliente.Checked, "VALORPRD", "VALORPRD"))
                    parameters.Add("Valor/Ton", IIf(RadCliente.Checked, "VALOR", "UNITARIO"))
                    parameters.Add("ValorFrete", IIf(RadCliente.Checked, "", "TOTAL"))

                    Funcoes.BindReport(Me.Page, ds, IIf(RadCliente.Checked, "Cr_RelPedidosEntreguesDeVendasClientes", "Cr_RelPedidosEntreguesDeVendas"), IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorioDados()
        Try
            If Funcoes.VerificaPermissao("PedidosEntreguesDeVendas", "RELATORIO") Then
                If Validar() Then
                    Dim ds As DataSet = getDataSet(True)

                    Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))


                    Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    If File.Exists(fileName) Then
                        File.Delete(fileName)
                    End If

                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)
                            'criando planilha títulos
                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("PEDIDOS ENTREGUES")

                            'criando linha com o cabeçalho da planilha
                            Dim rowIndex As Integer = 1
                            Dim columnIndex As Integer = 1

                            'criando linha que informa o nome da empresa e o cnpj
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa a cidade e o estado da empresa
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa o título do relatório
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELATÓRIO DE PEDIDOS ENTREGUES DE VENDAS")
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha que informa o período selecionado na página
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                            worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                            worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            rowIndex += 1

                            'criando linha com o cabeçalho da planilha
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            Next

                            'criando auto filtro na planilha
                            worksheet.Cells("A5:F" & rowIndex).AutoFilter = True

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using
                            rowIndex += 1

                            ' criando conteúdo da planilha com os dados do dataset
                            For Each row As DataRow In ds.Tables(0).Rows
                                columnIndex = 1
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    columnIndex += 1
                                Next

                                'formatando células datas
                                worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                                'formatando células numéricas
                                worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"
                                worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000000000_ ;[Red]-#,##0.0000000000"
                                worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                'aplicando formatação nas células do conteúdo
                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If
                                rowIndex += 1
                            Next

                            rowIndex += 1

                            'criando colunas de totalizadores
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(String.Format("G{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                            'criando colunas de totalizadores
                            worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Bold = True
                            worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                            worksheet.Cells(String.Format("N{0}", rowIndex)).Formula = String.Format("=SUM(N6:N{0})", rowIndex - 1)
                            worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            ''criando colunas de totalizadores
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Bold = True
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Formula = String.Format("=SUM(O6:O{0})", rowIndex - 1)
                            worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            ''criando colunas de totalizadores
                            worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Font.Bold = True
                            worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                            worksheet.Cells(String.Format("Q{0}", rowIndex)).Formula = String.Format("=SUM(Q6:Q{0})", rowIndex - 1)
                            worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'setando autofit nas células da planilha
                            worksheet.Cells.AutoFitColumns(0)

                            'congelando quinta linha (cabeçalho)
                            worksheet.View.FreezePanes(6, 1)

                            'salvando planilha do excel
                            package.Save()
                        End Using
                    End Using

                    Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PedidosEntreguesDeVendas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class