Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PesosDeChegada
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Expedicao)
                If Funcoes.VerificaPermissao("PesosDeChegada", "ACESSAR") Then
                    CarregarEmpresa()
                    CarregarGrupo()
                    Limpar()
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    ucConsultaPedidos.SetarHID(HID.Value)
                    ucPesoDeChegada.SetarHID(HID.Value)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página. É necessário ter as permissões ACESSAR, LEITURA, GRAVAR e RELATORIO no Processo PesosDeChegada", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objClientePesXChe" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClientePesXChe" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClientePesXChe" & HID.Value)
        ElseIf Not Session("objTransPesXChe" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objTransPesXChe" & HID.Value), [Lib].Negocio.Cliente))
            txtTransportador.Text = itemCliente.Text
            txtCodigoTrans.Value = itemCliente.Value
            Session.Remove("objTransPesXChe" & HID.Value)
        ElseIf Not Session("objPedidoSelecionadoPesXChe" & HID.Value) Is Nothing Then
            Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
            If CType(obj, [Lib].Negocio.Pedido).CodigoEmpresa <> Emp(0) Or CType(obj, [Lib].Negocio.Pedido).EnderecoEmpresa.ToString <> Emp(1) Then
                MsgBox(Me.Page, "Empresa do pedido é diferente da empresa selecionada.")
            ElseIf CType(obj, [Lib].Negocio.Pedido).CodigoCliente <> strCliente(0) Or CType(obj, [Lib].Negocio.Pedido).EnderecoCliente.ToString <> strCliente(1) Then
                MsgBox(Me.Page, "Cliente do pedido é diferente do cliente selecionado!")
            Else
                txtPedido.Text = CType(obj, [Lib].Negocio.Pedido).Codigo
                ddlGrupo.Enabled = False
                ddlProduto.Enabled = False
            End If
            Session.Remove("objPedidoSelecionadoPesXChe" & HID.Value)
        End If
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarGrupo()
        ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub CarregarProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & ddlGrupo.SelectedValue & "' AND Agrupar = 'N' ", True)
    End Sub

    Public Sub CarregarNotasFiscais(Optional ByVal emiteMsg As Boolean = False)
        Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim sql As String = String.Empty

        If chkViculados.Checked Then
            sql = "SELECT DISTINCT                                                                           " & vbCrLf & _
                "	      NF.EntradaSaida_Id, NF.Nota_Id, NF.Serie_Id,  " & vbCrLf & _
                "         NF.Cliente_Id, NF.EndCliente_Id, Cli.Nome, " & vbCrLf & _
                "         NF.Operacao, NF.SubOperacao, NF.Movimento AS Data," & vbCrLf & _
                "         NFxI.QuantidadeFiscal AS PesoFiscal, NFxI.Valor,  " & vbCrLf & _
                "         NxD.Movimento, NxD.PesoBruto, NxD.Desconto,  " & vbCrLf & _
                "         NxD.PesoLiquido, isnull(NxT.Placa,'') AS Placa, " & vbCrLf & _
                "         ISNULL(NxD.TarifaFrete,0) AS TarifaFrete, ISNULL(NxD.PesoBruto,0) * ISNULL(NxD.TarifaFrete,0) /1000 AS ValorFrete, " & vbCrLf

            If FinanceiroNovo Then
                sql &= "         CASE WHEN ISNULL(FFxT.QtdDeTitulos,0) > 0 THEN 'S' ELSE 'N' END Financeiro " & vbCrLf
            Else
                sql &= "         CASE WHEN ISNULL(FFxT.QtdDeTitulos,0) > 0 THEN 'S' ELSE 'N' END Financeiro " & vbCrLf
            End If

            sql &= getSqlFrom()

            sql &= " INNER JOIN NotasXDestinos NxD                                                 " & vbCrLf & _
                   "    ON NxD.Empresa_Id       = NFxI.Empresa_Id      " & vbCrLf & _
                   "   AND NxD.EndEmpresa_Id   = NFxI.EndEmpresa_Id    " & vbCrLf & _
                   "   AND NxD.Cliente_Id      = NFxI.Cliente_Id       " & vbCrLf & _
                   "   AND NxD.EndCliente_Id   = NFxI.EndCliente_Id    " & vbCrLf & _
                   "   AND NxD.EntradaSaida_Id = NFxI.EntradaSaida_Id  " & vbCrLf & _
                   "   AND NxD.Serie_Id        = NFxI.Serie_Id         " & vbCrLf & _
                   "   AND NxD.Nota_Id         = NFxI.Nota_Id          " & vbCrLf & _
                   "  -- AND NxD.Produto_Id      = NFxI.Produto_Id       " & vbCrLf & _
                   "  LEFT JOIN NotasXNotas NxN                                                     " & vbCrLf & _
                   "	ON NxN.OrigemEmpresa_Id      = NF.Empresa_Id         " & vbCrLf & _
                   "   AND NxN.OrigemEndEmpresa_Id   = NF.EndEmpresa_Id       " & vbCrLf & _
                   "   AND NxN.OrigemCliente_Id      = NF.Cliente_Id          " & vbCrLf & _
                   "   AND NxN.OrigemEndCliente_Id   = NF.EndCliente_Id       " & vbCrLf & _
                   "   AND NxN.OrigemEntradaSaida_Id = NF.EntradaSaida_Id     " & vbCrLf & _
                   "   AND NxN.OrigemSerie_Id        = NF.Serie_Id            " & vbCrLf & _
                   "   AND NxN.OrigemNota_Id         = NF.Nota_Id             " & vbCrLf & _
                   "  LEFT JOIN NotasFiscais CTRC                                               " & vbCrLf & _
                   "	ON NxN.Empresa_Id      = CTRC.Empresa_Id                       " & vbCrLf & _
                   "   AND NxN.EndEmpresa_Id   = CTRC.EndEmpresa_Id                     " & vbCrLf & _
                   "   AND NxN.Cliente_Id      = CTRC.Cliente_Id                        " & vbCrLf & _
                   "   AND NxN.EndCliente_Id   = CTRC.EndCliente_Id                     " & vbCrLf & _
                   "   AND NxN.EntradaSaida_Id = CTRC.EntradaSaida_Id                   " & vbCrLf & _
                   "   AND NxN.Serie_Id        = CTRC.Serie_Id                          " & vbCrLf & _
                   "   AND NxN.Nota_Id         = CTRC.Nota_Id                           " & vbCrLf & _
                   "  LEFT JOIN FaturasDeFretesXItens FFxI                                           " & vbCrLf & _
                   "	ON FFxI.Empresa_Id      = CTRC.Empresa_Id             " & vbCrLf & _
                   "   AND FFxI.EndEmpresa_Id   = CTRC.EndEmpresa_Id           " & vbCrLf & _
                   "   AND FFxI.Cliente_Id      = CTRC.Cliente_Id              " & vbCrLf & _
                   "   AND FFxI.EndCliente_Id   = CTRC.EndCliente_Id           " & vbCrLf & _
                   "   AND FFxI.EntradaSaida_Id = CTRC.EntradaSaida_Id         " & vbCrLf & _
                   "   AND FFxI.Serie_Id        = CTRC.Serie_Id                " & vbCrLf & _
                   "   AND FFxI.Nota_Id         = CTRC.Nota_Id                 " & vbCrLf & _
                   "  LEFT JOIN FaturasDeFretes  FF" & vbCrLf & _
                   "    ON FF.Fatura_Id = FFxI.Fatura_Id          " & vbCrLf

            If FinanceiroNovo Then
                '-- TENDO TITULO NÃO PODE SER DELETADO PORQUE JÁ FOI PROGRAMADO - FURLAN - 25/03/2015 **** Atualizado para nova estrutura 2015.10.08 - Cleberson
                sql &= "  LEFT JOIN ( SELECT COUNT(*) QtdDeTitulos, FFxTit.Empresa_Id, FFxTit.EndEmpresa_Id, FFxTit.Conveniado_Id, FFxTit.EndConveniado_Id, FFxTit.Fatura_Id   " & vbCrLf & _
                       "               FROM FaturaDeFreteXTitulo  FFxTit" & vbCrLf & _
                       "               JOIN Titulos T" & vbCrLf & _
                       "                 ON T.Titulo_Id = FFxTit.Titulo_Id" & vbCrLf & _
                       "                AND T.Situacao_Id = 1" & vbCrLf & _
                       "                AND T.Provisao_Id <> 3" & vbCrLf & _
                       "              GROUP BY FFxTit.Empresa_Id, FFxTit.EndEmpresa_Id, FFxTit.Conveniado_Id, FFxTit.EndConveniado_Id, FFxTit.Fatura_Id   " & vbCrLf & _
                       "             ) AS FFxT" & vbCrLf
            Else
                '-- TENDO TITULO NÃO PODE SER DELETADO PORQUE JÁ FOI PROGRAMADO - FURLAN - 25/03/2015 *** Atualizado para nova estrutura 2015.10.08 - Cleberson
                sql &= "  LEFT JOIN ( SELECT COUNT(*) QtdDeTitulos, FFxTit.Empresa_Id, FFxTit.EndEmpresa_Id, FFxTit.Conveniado_Id, FFxTit.EndConveniado_Id, FFxTit.Fatura_Id   " & vbCrLf &
                       "                FROM FaturaDeFreteXTitulo  FFxTit" & vbCrLf &
                       "                LEFT JOIN ContasAPagar CP" & vbCrLf &
                       "                  ON CP.Registro_Id = FFxTit.Titulo_Id" & vbCrLf &
                       "                 AND CP.Situacao = 1" & vbCrLf &
                       "                 AND CP.Provisao <> 3" & vbCrLf &
                       "                LEFT JOIN ContasAReceber CR" & vbCrLf &
                       "                  ON CR.Registro_Id = FFxTit.Titulo_Id" & vbCrLf &
                       "                 AND CR.Situacao = 1" & vbCrLf &
                       "                 AND CR.Provisao <> 3" & vbCrLf &
                       "               GROUP BY FFxTit.Empresa_Id, FFxTit.EndEmpresa_Id, FFxTit.Conveniado_Id, FFxTit.EndConveniado_Id, FFxTit.Fatura_Id   " & vbCrLf &
                       "            ) AS FFxT" & vbCrLf
            End If

            sql &= "    ON FFxI.EmpresaPagadora_Id = FFxT.Empresa_Id" & vbCrLf & _
                   "   AND FFxI.EndEmpresaPagadora_Id = FFxT.EndEmpresa_Id" & vbCrLf & _
                   "   AND FFxI.Conveniado_Id = FFxT.Conveniado_Id" & vbCrLf & _
                   "   AND FFxI.EndConveniado_Id = FFxT.EndConveniado_Id" & vbCrLf & _
                   "   AND FFxI.Fatura_Id = FFxT.Fatura_Id" & vbCrLf & _
                   " WHERE 1=1" & vbCrLf & _
                   "   AND NF.Situacao = 1" & vbCrLf & _
                   "   AND NF.TipoDeDocumento = 1" & vbCrLf & _
                   "   AND (NF.Empresa_Id      = '" & Emp(0) & "')" & vbCrLf & _
                   "   AND (NF.EndEmpresa_Id   = " & Emp(1) & ")" & vbCrLf & _
                   "   AND (SOP.QuantidadeFisico = 'S' OR SOP.CLASSE = '" & eClassesOperacoes.VENDASAORDEM.ToString & "') " & vbCrLf

            If Not String.IsNullOrWhiteSpace(txtNotaFiscal.Text) Then
                sql &= "  AND (NF.Nota_Id = '" & txtNotaFiscal.Text.Trim() & "') " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                sql &= "  AND (NF.Cliente_Id    = '" & strCliente(0) & "')" & vbCrLf & _
                       "  AND (NF.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtTransportador.Text) Then
                Dim strTrans() As String = txtCodigoTrans.Value.Split("-")
                sql &= "  AND (NxT.Proprietario    = '" & strTrans(0) & "')" & vbCrLf & _
                       "  AND (NxT.EndProprietario = " & strTrans(1) & ")" & vbCrLf
            End If

            sql &= "  AND (NF.EntradaSaida_Id = '" & IIf(rdSaida.Checked, "S", "E") & "')" & vbCrLf

            If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                sql &= "  AND (NF.Pedido = " & txtPedido.Text & ")" & vbCrLf
            End If

            If ddlProduto.SelectedIndex > 0 Then
                sql &= "  AND (NFxI.Produto_Id = '" & ddlProduto.SelectedValue & "')" & vbCrLf
            End If

            If String.IsNullOrWhiteSpace(txtNotaFiscal.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                sql &= "  AND (NF.Movimento BETWEEN '" & txtDataInicial.Text.Trim().ToSqlDate() & "' AND '" & txtDataFinal.Text.Trim().ToSqlDate() & "')"
            End If

        Else
            sql = "SELECT NF.EntradaSaida_Id, NF.Nota_Id, NF.Serie_Id, " & vbCrLf & _
                  "       NF.Cliente_Id, NF.EndCliente_Id, Cli.Nome, " & vbCrLf & _
                  "       NF.Operacao, NF.SubOperacao, NF.Movimento AS Data, " & vbCrLf & _
                  "       NFxI.QuantidadeFiscal AS PesoFiscal, NFxI.Valor, 'N' as Financeiro,  " & vbCrLf & _
                  "       GETDATE() Movimento, 0.0 PesoBruto, 0.0 Desconto, 0.0 PesoLiquido, isnull(NxT.Placa,'') AS Placa " & vbCrLf

            sql &= getSqlFrom()

            sql &= " WHERE 1=1 " & vbCrLf & _
                   "   And (NF.Situacao = 1)" & vbCrLf & _
                   "   AND (NF.TipoDeDocumento = 1) " & vbCrLf & _
                   "   AND (NF.Empresa_Id      = '" & Emp(0) & "')" & vbCrLf & _
                   "   AND (NF.EndEmpresa_Id   = " & Emp(1) & ")" & vbCrLf

            If Not String.IsNullOrWhiteSpace(txtNotaFiscal.Text) Then
                sql &= "  AND (NF.Nota_Id = '" & txtNotaFiscal.Text.Trim() & "') " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                sql &= "  AND (NF.Cliente_Id    = '" & strCliente(0) & "')" & vbCrLf & _
                       "  AND (NF.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtTransportador.Text) Then
                Dim strTrans() As String = txtCodigoTrans.Value.Split("-")
                sql &= "  AND (NxT.Proprietario    = '" & strTrans(0) & "')" & vbCrLf & _
                       "  AND (NxT.EndProprietario = " & strTrans(1) & ")" & vbCrLf
            End If

            sql &= "  AND (NF.EntradaSaida_Id = '" & IIf(rdSaida.Checked, "S", "E") & "')" & vbCrLf

            sql &= "  AND (NF.CIFFOB = '" & IIf(rdSaida.Checked, "CIF", "FOB") & "')" & vbCrLf

            sql &= "  AND (SOP.FinalidadeDaNota NOT IN(2))" & vbCrLf

            If rdSaida.Checked Then
                sql &= "   AND (SOP.QuantidadeFisico = 'S' OR SOP.CLASSE = 'CONTAEORDEM' OR SOP.CLASSE = '" & eClassesOperacoes.VENDASAORDEM.ToString & "') " & vbCrLf
            Else
                sql &= "   AND (SOP.QuantidadeFisico = 'S' OR SOP.CLASSE = 'CONTAEORDEM' OR SOP.CLASSE = '" & eClassesOperacoes.COMPRASAORDEM.ToString & "') " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
                sql &= "  AND (NF.Pedido = " & txtPedido.Text & ")" & vbCrLf
            End If

            If ddlProduto.SelectedIndex > 0 Then
                sql &= "  AND (NFxI.Produto_Id = '" & ddlProduto.SelectedValue & "')" & vbCrLf
            End If

            If String.IsNullOrWhiteSpace(txtNotaFiscal.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                sql &= "  AND (NF.Movimento BETWEEN '" & txtDataInicial.Text.Trim().ToSqlDate() & "' AND '" & txtDataFinal.Text.Trim().ToSqlDate() & "')" & vbCrLf
            End If

            sql &= "  AND NOT EXISTS(SELECT NxD.Nota_Id " & vbCrLf & _
                   "                   FROM NotasXDestinos NxD " & vbCrLf & _
                   "                  WHERE NxD.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                   "                    AND NxD.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                   "                    AND NxD.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                   "                    AND NxD.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                   "                    AND NxD.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                   "                    AND NxD.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                   "                    AND NxD.Nota_Id         = NFxI.Nota_Id) " & vbCrLf & _
                   "                    --AND NxD.Produto_Id      = NFxI.Produto_Id " & vbCrLf & _
                   "ORDER BY NF.Movimento, NF.Nota_Id " & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasFiscais")
        'NFD 02
        If ds.Tables(0).Rows.Count = 0 AndAlso emiteMsg Then MsgBox(Me.Page, "Nenhum Registro Encontrado!")
        
        gridNF.DataSource = ds
        gridNF.DataBind()
    End Sub

    Private Function getSqlFrom()
        Dim Sql As String = String.Empty

        Sql &= "  FROM NotasFiscais NF " & vbCrLf & _
               " INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
               "    ON NF.Empresa_Id       = NFxI.Empresa_Id " & vbCrLf & _
               "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
               "   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
               "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
               "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
               "   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
               "   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
               " INNER JOIN Produtos P " & vbCrLf & _
               "    ON NFxI.Produto_Id = P.Produto_Id " & vbCrLf & _
               " INNER JOIN SubOperacoes Sop" & vbCrLf & _
               "    ON NF.Operacao = Sop.Operacao_Id" & vbCrLf & _
               "   AND NF.SubOperacao = Sop.SubOperacoes_Id" & vbCrLf & _
               " INNER JOIN Clientes Cli " & vbCrLf & _
               "    ON NF.Cliente_Id    = Cli.Cliente_Id " & vbCrLf & _
               "   AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf & _
               "  LEFT JOIN NotasFiscaisXTransportadores NxT" & vbCrLf & _
               "	ON NF.Empresa_Id       = NxT.Empresa_Id            " & vbCrLf & _
               "   AND NF.EndEmpresa_Id   = NxT.EndEmpresa_Id     " & vbCrLf & _
               "   AND NF.Cliente_Id      = NxT.Cliente_Id           " & vbCrLf & _
               "   AND NF.EndCliente_Id   = NxT.EndCliente_Id     " & vbCrLf & _
               "   AND NF.EntradaSaida_Id = NxT.EntradaSaida_Id " & vbCrLf & _
               "   AND NF.Serie_Id        = NxT.Serie_Id               " & vbCrLf & _
               "   AND NF.Nota_Id         = NxT.Nota_Id                 " & vbCrLf
        Return Sql
    End Function

    Private Sub Limpar()
        Session.Remove("objClientePesXChe" & HID.Value)
        Session.Remove("objPedidoSelecionadoPesXChe" & HID.Value)
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        txtTransportador.Text = ""
        txtCodigoTrans.Value = ""
        ddlGrupo.Enabled = True
        ddlProduto.Enabled = True
        txtPedido.Text = ""
        btnPedido.Enabled = True
        ddlGrupo.SelectedIndex = 0
        ddlProduto.Items.Clear()
        Funcoes.VerificaEmpresa(ddlEmpresa)
        txtDataInicial.Text = Now().ToString("dd/MM/yyyy")
        txtDataFinal.Text = Now().ToString("dd/MM/yyyy")
        gridNF.DataSource = Nothing
        gridNF.DataBind()
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período: " & txtDataInicial.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à: " & txtDataFinal.Text & vbCrLf
        End If
        param &= IIf(rdEntrada.Checked, "E/S: Entrada", "E/S: Saída")

        Return param
    End Function

    Protected Sub btnTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
            Popup.ConsultaDeClientes(Me.Page, "objTransPesXChe" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClientePesXChe" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
                MsgBox(Me.Page, "Cliente não foi selecionado.")
            Else
                HttpContext.Current.Session("ssCampo" & HID.Value) = "CargaPedidos"
                Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                Dim parameters As New Dictionary(Of String, Object)
                parameters("unidade") = ""
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                Popup.ConsultaDePedidos(Me.Page, "objPedidoSelecionadoPesXChe" & HID.Value)
                ucConsultaPedidos.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlGrupo.SelectedIndex > 0 Then
                CarregarProduto()
                btnPedido.Enabled = False
            Else
                ddlProduto.Items.Clear()
                btnPedido.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkSelecionarNF_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnkPedido As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnkPedido.NamingContainer, GridViewRow)
            Dim Emp() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal()
            objNotaFiscal.CodigoEmpresa = Emp(0)
            objNotaFiscal.EnderecoEmpresa = Emp(1)
            objNotaFiscal.CodigoCliente = gridNF.Rows(row.RowIndex).Cells(4).Text()
            objNotaFiscal.EnderecoCliente = gridNF.Rows(row.RowIndex).Cells(5).Text()
            objNotaFiscal.EntradaSaida = IIf(gridNF.Rows(row.RowIndex).Cells(1).Text() = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
            objNotaFiscal.Serie = gridNF.Rows(row.RowIndex).Cells(3).Text()
            objNotaFiscal.Codigo = gridNF.Rows(row.RowIndex).Cells(2).Text()
            objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

            Dim Sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id " & vbCrLf &
                  "  FROM NotasXNotas " & vbCrLf &
                  " WHERE OrigemEmpresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                  "   AND OrigemEndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                  "   AND OrigemCliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                  "   AND OrigemEndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                  "   AND OrigemEntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "   AND OrigemSerie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                  "   AND OrigemNota_Id         = " & objNotaFiscal.Codigo & vbCrLf

            Dim dsCte As New DataSet
            dsCte = Banco.ConsultaDataSet(Sql, "NotasFiscaisXCte")

            If Not dsCte Is Nothing AndAlso dsCte.Tables(0).Rows.Count > 0 Then
                Dim objCte As New [Lib].Negocio.NotaFiscal()
                objCte.CodigoEmpresa = dsCte.Tables(0).Rows(0).Item("Empresa_Id")
                objCte.EnderecoEmpresa = dsCte.Tables(0).Rows(0).Item("EndEmpresa_Id")
                objCte.CodigoCliente = dsCte.Tables(0).Rows(0).Item("Cliente_Id")
                objCte.EnderecoCliente = dsCte.Tables(0).Rows(0).Item("EndCliente_Id")
                objCte.EntradaSaida = IIf(dsCte.Tables(0).Rows(0).Item("EntradaSaida_Id") = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
                objCte.Serie = dsCte.Tables(0).Rows(0).Item("Serie_Id")
                objCte.Codigo = dsCte.Tables(0).Rows(0).Item("Nota_Id")
                objCte = New [Lib].Negocio.NotaFiscal(objCte)

                If objCte.CodigoTipoDeDocumento = 2 Or objCte.CodigoTipoDeDocumento = 57 Then
                    Sql = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id " & vbCrLf &
                          "  FROM NotasXNotas " & vbCrLf &
                          " WHERE OrigemEmpresa_Id      = '" & objCte.CodigoEmpresa & "'" & vbCrLf &
                          "   AND OrigemEndEmpresa_Id   = " & objCte.EnderecoEmpresa & vbCrLf &
                          "   AND OrigemCliente_Id      = '" & objCte.CodigoCliente & "'" & vbCrLf &
                          "   AND OrigemEndCliente_Id   = " & objCte.EnderecoCliente & vbCrLf &
                          "   AND OrigemEntradaSaida_Id = '" & objCte.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   AND OrigemSerie_Id        = '" & objCte.Serie & "'" & vbCrLf &
                          "   AND OrigemNota_Id         = " & objCte.Codigo & vbCrLf &
                          "   AND Serie_Id              = '10'" & vbCrLf

                    Dim dsCtro As New DataSet
                    dsCtro = Banco.ConsultaDataSet(Sql, "CteXCtro")

                    If Not dsCtro Is Nothing AndAlso dsCtro.Tables(0).Rows.Count > 0 Then
                        Dim objCtro As New [Lib].Negocio.NotaFiscal()
                        objCtro.CodigoEmpresa = dsCtro.Tables(0).Rows(0).Item("Empresa_Id")
                        objCtro.EnderecoEmpresa = dsCtro.Tables(0).Rows(0).Item("EndEmpresa_Id")
                        objCtro.CodigoCliente = dsCtro.Tables(0).Rows(0).Item("Cliente_Id")
                        objCtro.EnderecoCliente = dsCtro.Tables(0).Rows(0).Item("EndCliente_Id")
                        objCtro.EntradaSaida = IIf(dsCtro.Tables(0).Rows(0).Item("EntradaSaida_Id") = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
                        objCtro.Serie = dsCtro.Tables(0).Rows(0).Item("Serie_Id")
                        objCtro.Codigo = dsCtro.Tables(0).Rows(0).Item("Nota_Id")
                        objCtro = New [Lib].Negocio.NotaFiscal(objCtro)

                        If objCtro.CodigoTipoDeDocumento = 4 Then
                            Sql = "SELECT Titulo_Id " & vbCrLf &
                                  "  FROM NotaFiscalXTitulo " & vbCrLf &
                                  " WHERE Empresa_Id      = '" & objCtro.CodigoEmpresa & "'" & vbCrLf &
                                  "   AND EndEmpresa_Id   = " & objCtro.EnderecoEmpresa & vbCrLf &
                                  "   AND Cliente_Id      = '" & objCtro.CodigoCliente & "'" & vbCrLf &
                                  "   AND EndCliente_Id   = " & objCtro.EnderecoCliente & vbCrLf &
                                  "   AND EntradaSaida_Id = '" & objCtro.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                  "   AND Serie_Id        = '" & objCtro.Serie & "'" & vbCrLf &
                                  "   AND Nota_Id         = " & objCtro.Codigo & vbCrLf

                            Dim dsTit As New DataSet
                            dsTit = Banco.ConsultaDataSet(Sql, "CtroXTit")

                            If Not dsTit Is Nothing AndAlso dsCtro.Tables(0).Rows.Count > 0 Then
                                Dim objTitulo As New [Lib].Negocio.Titulo(dsTit.Tables(0).Rows(0).Item("Titulo_Id"))

                                If objTitulo.CodigoProvisao = 1 Then
                                    MsgBox(Me.Page, "Vencimento do Saldo do Contrato Pamcard já está Baixado. Entre em contato com o Suporte.")
                                Else
                                    Session("objContratoDeFrete") = objCtro
                                    Session("objTituloContratoDeFrete") = objTitulo
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            Session("objNFPesoDeChegada") = objNotaFiscal
            ucPesoDeChegada.LimparCampos()
            ucPesoDeChegada.BindGridView(objNotaFiscal.Itens(0).CodigoProduto, objNotaFiscal.Itens(0).PesoFiscal)
            Popup.ConsultaDePesoDeChegada(Me.Page, "objPesoDeChegada" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgDelete_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgDelete As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgDelete.NamingContainer, GridViewRow)
            Dim Sqls As New ArrayList
            Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal()
            Dim ObjNota As New [Lib].Negocio.NotaFiscal()

            objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
            objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)

            objNotaFiscal.CodigoCliente = gridNF.Rows(row.RowIndex).Cells(4).Text()
            objNotaFiscal.EnderecoCliente = gridNF.Rows(row.RowIndex).Cells(5).Text()
            objNotaFiscal.EntradaSaida = IIf(gridNF.Rows(row.RowIndex).Cells(1).Text() = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
            objNotaFiscal.Serie = gridNF.Rows(row.RowIndex).Cells(3).Text()
            objNotaFiscal.Codigo = gridNF.Rows(row.RowIndex).Cells(2).Text()
            ObjNota = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

            For Each objAnalise As [Lib].Negocio.NotaFiscalXDestinoXDescontos In ObjNota.Itens(0).DescontosPesoDeChegada
                objAnalise.IUD = "D"
                objAnalise.SalvarSql(Sqls)
            Next

            objNotaFiscal.PesoDeChegada.IUD = "D"
            objNotaFiscal.PesoDeChegada.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

            MsgBox(Me.Page, "Dados excluídos com Sucesso.", eTitulo.Sucess)
            CarregarNotasFiscais()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("PesosDeChegada", "LEITURA") Then
                If ddlEmpresa.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Empresa não foi selecionada.")

                ElseIf chkViculados.Checked Then
                    'Selecionar
                    gridNF.Columns(0).Visible = False
                    'Movimento
                    gridNF.Columns(12).Visible = True
                    'Peso Bruto
                    gridNF.Columns(13).Visible = True
                    'Desconto
                    gridNF.Columns(14).Visible = True
                    'Peso Liquido
                    gridNF.Columns(15).Visible = True
                    'Placas
                    gridNF.Columns(17).Visible = False

                    'Tarifa do Frete
                    gridNF.Columns(18).Visible = True
                    'Valor Do Frete
                    gridNF.Columns(19).Visible = True

                    'Excluir lançamento
                    gridNF.Columns(20).Visible = True
                Else
                    'Selecionar
                    gridNF.Columns(0).Visible = True
                    'Movimento
                    gridNF.Columns(12).Visible = False
                    'Peso Bruto
                    gridNF.Columns(13).Visible = False
                    'Desconto
                    gridNF.Columns(14).Visible = False
                    'Peso Liquido
                    gridNF.Columns(15).Visible = False
                    'Placas
                    gridNF.Columns(17).Visible = True
                    'Tarifa do Frete
                    gridNF.Columns(18).Visible = False
                    'Valor Do Frete
                    gridNF.Columns(19).Visible = False
                    'Excluir Lançamento
                    gridNF.Columns(20).Visible = True
                End If
                CarregarNotasFiscais(True)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro(s).")
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

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("PesosDeChegada", "RELATORIO") Then

                Dim ds As DataSet = getDataSet()
                Dim param As New Dictionary(Of String, Object)()
                param.Add("Parametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_PesosDeChegada", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), param)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridNF_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridNF.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If e.Row.DataItem IsNot Nothing Then
                    Dim lblFinanceiro As Label = CType(e.Row.FindControl("lblFinanceiro"), Label)
                    Dim imgDelete As ImageButton = CType(e.Row.FindControl("imgDelete"), ImageButton)

                    If chkViculados.Checked Then
                        If lblFinanceiro.Text.ToUpper().Trim().Contains("N") Then
                            imgDelete.Visible = True
                        Else
                            imgDelete.Visible = False
                        End If
                    Else
                        imgDelete.Visible = False
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim Sql As String = " SELECT Consulta.Nota_Id as  Nota, " & vbCrLf &
                            "        Consulta.Serie_Id as  Serie, " & vbCrLf &
                            "        isnull(Consulta.Placa,'') as Placa," & vbCrLf &
                            "        Consulta.DataSaida," & vbCrLf &
                            "        Consulta.DataChegada," & vbCrLf &
                            "        Consulta.Operacao, " & vbCrLf &
                            "        Consulta.SubOperacao, " & vbCrLf &
                            "        Consulta.Origem, " & vbCrLf &
                            "        Consulta.Bruto, " & vbCrLf &
                            "        Consulta.Desconto, " & vbCrLf &
                            "        Consulta.Liquido," & vbCrLf &
                            "        Sum(ISNULL(CASE WHEN NFxDD.Analise_Id = 01 THEN NFxDD.Desconto END, 0)) AS Umidade," & vbCrLf &
                            "        Sum(ISNULL(CASE WHEN NFxDD.Analise_Id = 02 THEN NFxDD.Desconto END, 0)) AS Impureza, " & vbCrLf &
                            "        Sum(ISNULL(CASE WHEN NFxDD.Analise_Id = 03 THEN NFxDD.Desconto END, 0)) AS Avariado, " & vbCrLf &
                            "        Sum(ISNULL(CASE WHEN NFxDD.Analise_Id = 08 THEN NFxDD.Desconto END, 0)) AS Retencao" & vbCrLf &
                            "   FROM (" & vbCrLf &
                            "          SELECT NF.Empresa_Id, " & vbCrLf &
                            "                 NF.EndEmpresa_Id, " & vbCrLf &
                            "                 NF.Cliente_Id, " & vbCrLf &
                            "                 NF.EndCliente_Id, " & vbCrLf &
                            "                 NF.Nota_Id," & vbCrLf &
                            "                 NF.Serie_Id, " & vbCrLf &
                            "                 NF.EntradaSaida_Id, " & vbCrLf &
                            "                 NFxI.Produto_Id," & vbCrLf &
                            "                 NF.Movimento as DataSaida, " & vbCrLf &
                            "                 NFxD.Movimento as DataChegada," & vbCrLf &
                            "                 NF.Operacao, " & vbCrLf &
                            "                 NF.SubOperacao," & vbCrLf &
                            "                 NxT.Placa," & vbCrLf &
                            "                 NFxI.QuantidadeFisica AS Origem," & vbCrLf &
                            "                 ISNULL(NFxD.PesoBruto, 0) AS Bruto," & vbCrLf &
                            "                 ISNULL(NFxD.Desconto, 0) AS Desconto, " & vbCrLf &
                            "                 ISNULL(NFxD.PesoLiquido, 0) AS Liquido" & vbCrLf &
                            "            FROM NotasXDestinos  NFxD " & vbCrLf &
                            "           RIGHT OUTER JOIN NotasFiscais NF  " & vbCrLf &
                            "           INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf &
                            "              ON NF.Empresa_Id = NFxI.Empresa_Id  " & vbCrLf &
                            "             AND NF.EndEmpresa_Id = NFxI.EndEmpresa_Id " & vbCrLf &
                            "	          AND NF.Cliente_Id = NFxI.Cliente_Id " & vbCrLf &
                            "             AND NF.EndCliente_Id = NFxI.EndCliente_Id " & vbCrLf &
                            "             AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
                            "		      AND NF.Serie_Id = NFxI.Serie_Id " & vbCrLf &
                            "             AND NF.Nota_Id = NFxI.Nota_Id " & vbCrLf &
                            "              ON NFxD.Empresa_Id = NFxI.Empresa_Id " & vbCrLf &
                            "             AND NFxD.EndEmpresa_Id = NFxI.EndEmpresa_Id " & vbCrLf &
                            "             AND NFxD.Cliente_Id = NFxI.Cliente_Id " & vbCrLf &
                            "             AND NFxD.EndCliente_Id = NFxI.EndCliente_Id " & vbCrLf &
                            "             AND NFxD.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
                            "             AND NFxD.Serie_Id = NFxI.Serie_Id " & vbCrLf &
                            "             AND NFxD.Nota_Id = NFxI.Nota_Id " & vbCrLf &
                            "             --AND NFxD.Produto_Id = NFxI.Produto_Id " & vbCrLf &
                            "            LEFT JOIN NotasFiscaisXTransportadores NxT " & vbCrLf &
                            "              ON NxT.Empresa_Id      = NF.Empresa_Id               " & vbCrLf &
                            "             AND NxT.EndEmpresa_Id   = NF.EndEmpresa_Id            " & vbCrLf &
                            "             AND NxT.Cliente_Id      = NF.Cliente_Id               " & vbCrLf &
                            "             AND NxT.EndCliente_Id   = NF.EndCliente_Id            " & vbCrLf &
                            "             AND NxT.EntradaSaida_Id = NF.EntradaSaida_Id          " & vbCrLf &
                            "             AND NxT.Serie_Id        = NF.Serie_Id                 " & vbCrLf &
                            "             AND NxT.Nota_Id         = NF.Nota_Id                  " & vbCrLf &
                            "             LEFT Join NotasFiscaisXRomaneios NxR " & vbCrLf &
                            "              ON NxR.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
                            "             AND NxR.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
                            "             AND NxR.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
                            "             AND NxR.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
                            "             AND NxR.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                            "             AND NxR.Serie_Id        = NF.Serie_Id " & vbCrLf &
                            "             AND NxR.Nota_Id         = NF.Nota_Id " & vbCrLf &
                            "            LEFT join RomaneiosXPesagens RxP " & vbCrLf &
                            "              ON NxR.Empresa_Id    = RxP.Empresa_Id " & vbCrLf &
                            "             AND NxR.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf &
                            "             AND NxR.Romaneio_Id   = RxP.Romaneio_Id " & vbCrLf &
                            "            Left join Pesagem " & vbCrLf &
                            "              ON RxP.Empresa_Id      = Pesagem.Empresa_Id " & vbCrLf &
                            "             AND RxP.EndEmpresa_Id   = Pesagem.EndEmpresa_Id " & vbCrLf &
                            "             AND RxP.Pesagem_Id     = Pesagem.Pesagem_Id " & vbCrLf &
                            "             AND Pesagem.Sequencia_Id = 0 " & vbCrLf &
                            "           INNER JOIN Produtos prd " & vbCrLf &
                            "              ON prd.Produto_id = NFxI.Produto_Id " & vbCrLf &
                            "           INNER JOIN SubOperacoes Sop " & vbCrLf &
                            "              ON NF.Operacao    = Sop.Operacao_Id " & vbCrLf &
                            "             AND NF.SubOperacao = Sop.SubOperacoes_Id " & vbCrLf &
                            "           WHERE (NF.Situacao = 1) " & vbCrLf &
                            "             AND (NF.TipoDeDocumento = 1)  " & vbCrLf &
                            "             AND (prd.Agrupar = 'N') " & vbCrLf &
                            "             AND (NF.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "') " & vbCrLf &
                            "             AND (NF.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & ")" & vbCrLf

        If rdEntrada.Checked = True Then
            Sql &= "             And (NF.EntradaSaida_Id = 'E') " & vbCrLf
        Else
            Sql &= "             And (NF.EntradaSaida_Id = 'S') " & vbCrLf
        End If

        If txtPedido.Text.Length > 0 Then
            Sql &= "             And (NF.Pedido= " & txtPedido.Text & ") " & vbCrLf
        End If

        If ddlProduto.SelectedIndex > 0 Then
            Sql &= "             And (NFxI.Produto_Id = '" & ddlProduto.SelectedValue & "') " & vbCrLf
        End If

        If txtCodigoCliente.Value.ToString.Length > 0 Then
            Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
            Sql &= "             And  (NF.Cliente_Id = '" & strCliente(0) & "') AND (NF.EndCliente_Id = " & strCliente(1) & ") " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtTransportador.Text) Then
            Dim strTrans() As String = txtCodigoTrans.Value.Split("-")
            Sql &= "             AND (NxT.Proprietario    = '" & strTrans(0) & "')" & vbCrLf &
                   "             AND (NxT.EndProprietario = " & strTrans(1) & ")" & vbCrLf
        End If

        Sql &= "             AND (Sop.FinalidadeDaNota NOT IN(2))" & vbCrLf

        If chkViculados.Checked Then
            Sql &= "             And (NOT NFxD.Movimento IS NULL) " & vbCrLf
        Else
            Sql &= "             And (NFxD.Movimento IS NULL) " & vbCrLf
        End If

        Sql &= "             AND  (NF.Movimento between  '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf &
               "             AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
               "        GROUP BY NF.Empresa_Id, " & vbCrLf &
               "		         NF.EndEmpresa_Id, " & vbCrLf &
               "		         NF.Cliente_Id, " & vbCrLf &
               "		         NF.EndCliente_Id, " & vbCrLf &
               "		         NF.Serie_Id, " & vbCrLf &
               "                 NF.EntradaSaida_Id," & vbCrLf &
               "		         NFxI.Produto_Id, " & vbCrLf &
               "		         NF.Movimento," & vbCrLf &
               "		         NFxD.Movimento, " & vbCrLf &
               "		         NF.Nota_Id, " & vbCrLf &
               "		         NF.Operacao, " & vbCrLf &
               "                 NF.SubOperacao, " & vbCrLf &
               "                 NxT.Placa," & vbCrLf &
               "		         NFxI.QuantidadeFisica, " & vbCrLf &
               "		         NFxD.PesoBruto," & vbCrLf &
               "		         NFxD.Desconto, " & vbCrLf &
               "                 NFxD.PesoLiquido) AS Consulta " & vbCrLf &
               "		    LEFT OUTER JOIN  NotasXDestinosXDescontos NFxDD " & vbCrLf &
               "              ON Consulta.Empresa_Id COLLATE Latin1_General_CI_AI = NFxDD.Empresa_Id  " & vbCrLf &
               "			 AND Consulta.EndEmpresa_Id = NFxDD.EndEmpresa_Id  " & vbCrLf &
               "			 AND Consulta.Cliente_Id COLLATE Latin1_General_CI_AI = NFxDD.Cliente_Id  " & vbCrLf &
               "			 AND Consulta.EndCliente_Id = NFxDD.EndCliente_Id AND Consulta.Nota_Id = NFxDD.Nota_Id  " & vbCrLf &
               "			 AND Consulta.Serie_Id COLLATE Latin1_General_CI_AI = NFxDD.Serie_Id  " & vbCrLf &
               "			 AND Consulta.EntradaSaida_Id COLLATE Latin1_General_CI_AI = NFxDD.EntradaSaida_Id  " & vbCrLf &
               "			 AND Consulta.Produto_Id COLLATE Latin1_General_CI_AI  = NFxDD.Produto_Id " & vbCrLf &
               " GROUP BY  Consulta.Nota_Id, Consulta.Serie_Id, Consulta.Placa, Consulta.DataSaida, Consulta.DataChegada, Consulta.Operacao, Consulta.SubOperacao, " & vbCrLf &
               "           Consulta.Origem, Consulta.Bruto, Consulta.Desconto, Consulta.Liquido" & vbCrLf
        Return Banco.ConsultaDataSet(Sql, "PesosDeChegada") '#NFD 03
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PesosDeChegada")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class