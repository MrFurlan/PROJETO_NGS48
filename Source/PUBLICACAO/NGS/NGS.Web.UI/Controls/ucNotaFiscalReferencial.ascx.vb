Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucNotaFiscalReferencial
    Inherits BaseUserControl

#Region "Properties"

    Private Property dsNotasFiscais() As DataSet
        Get
            Return Session("objDataSetNotasFiscais" + HID.Value)
        End Get
        Set(ByVal value As DataSet)
            Session("objDataSetNotasFiscais" + HID.Value) = value
        End Set
    End Property

    Private Property ListaDeNotasFiscais() As List(Of [Lib].Negocio.NotaFiscal)
        Get
            'If Session("objListaDeNotasFiscais" + HID.Value) Is Nothing Then
            '    Session("objListaDeNotasFiscais" + HID.Value) = New List(Of [Lib].Negocio.NotaFiscal)
            'End If
            Return Session("objListaDeNotasFiscais" + HID.Value)
        End Get
        Set(ByVal value As List(Of [Lib].Negocio.NotaFiscal))
            Session("objListaDeNotasFiscais" + HID.Value) = value
        End Set
    End Property

    Private Property QuantidadeFiscalNF() As Decimal
        Get
            Return Session("QuantidadeFiscalNF" + HID.Value)
        End Get
        Set(ByVal value As Decimal)
            Session("QuantidadeFiscalNF" + HID.Value) = value
        End Set
    End Property

    Private Property ValorNFOrigem() As Decimal
        Get
            Return Session("ValorNFOrigem" + HID.Value)
        End Get
        Set(ByVal value As Decimal)
            Session("ValorNFOrigem" + HID.Value) = value
        End Set
    End Property

    Private Property TipoReferencial() As eTipoReferencial
        Get
            Return CType(Session("TipoReferencial" + HID.Value), eTipoReferencial)
        End Get
        Set(ByVal value As eTipoReferencial)
            Session("TipoReferencial" + HID.Value) = CType(value, eTipoReferencial)
        End Set
    End Property

#End Region

#Region "Methods"

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtPesoTotalSelecionado.Text = 0
        txtValorTotalSelecionado.Text = 0

        If Not dsNotasFiscais Is Nothing Then
            ComporNotasFiscais(dsNotasFiscais)
        Else
            grdNFeSaldo.DataSource = New List(Of Object)
        End If
        grdNFeSaldo.DataBind()

    End Sub

    Public Sub ConsultarNotas(ByVal Parametros As Dictionary(Of String, Object))
        If Parametros.ContainsKey("QuantidadeFiscalNF") Then
            Me.QuantidadeFiscalNF = Parametros("QuantidadeFiscalNF")
        Else
            Me.QuantidadeFiscalNF = 0
        End If

        If Parametros.ContainsKey("TipoReferencial") Then
            Me.TipoReferencial = Parametros("TipoReferencial")
        Else
            Me.TipoReferencial = eTipoReferencial.CTE
        End If

        If Parametros.ContainsKey("ValorNFOrigem") Then
            Me.ValorNFOrigem = Parametros("ValorNFOrigem")
        Else
            Me.ValorNFOrigem = 0
        End If

        dsNotasFiscais = Banco.ConsultaDataSet(getSqlSaldo(Parametros), "NotasFiscais")

        ComporNotasFiscais(dsNotasFiscais)

    End Sub

    Private Sub ComporNotasFiscais(ds As DataSet)
        grdNFeSaldo.Visible = True

        Dim objNf As NotaFiscal

        If ListaDeNotasFiscais Is Nothing Then
            ListaDeNotasFiscais = New List(Of [Lib].Negocio.NotaFiscal)
        Else
            ListaDeNotasFiscais.Clear()
        End If

        For Each nf In ds.Tables(0).Rows
            objNf = New NotaFiscal()
            objNf.CodigoEmpresa = nf("Empresa_id")
            objNf.EnderecoEmpresa = nf("EndEmpresa_id")
            objNf.CodigoCliente = nf("Cliente_id")
            objNf.EnderecoCliente = nf("EndCliente_id")
            objNf.EntradaSaida = IIf(nf("EntradaSaida_id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objNf.Codigo = nf("Nota_id")
            objNf.Serie = nf("Serie_Id")
            objNf = New NotaFiscal(objNf)
            ListaDeNotasFiscais.Add(objNf)
        Next

        If Me.TipoReferencial = eTipoReferencial.RPA Then
            lblTipoReferencial.Text = "Notas Referenciais de Recibo de Pagamento a Autônomo (RPA)"
            divPaineltotais.Style.Add("display", "none")
            txtValorRPA.Text = Me.ValorNFOrigem.ToString("N2")
        ElseIf Me.TipoReferencial = eTipoReferencial.EXP Then
            lblTipoReferencial.Text = "Notas Referenciais de Exportação"
            divValorNFOrigem.Style.Add("display", "none")
        ElseIf Me.TipoReferencial = eTipoReferencial.CTE Then
            lblTipoReferencial.Text = "Notas Referenciais de Conhecimento de Transporte"
            divValorNFOrigem.Style.Add("display", "none")
        End If

        grdNFeSaldo.DataSource = dsNotasFiscais
        grdNFeSaldo.DataBind()

        txtSaldoTotal.Text = calculaSaldoTotal()
        txtPesoTotal.Text = calculaPesoAjustavel()
        txtTotalValorAtualizado.Text = calculaValorTotalAtualizado()
    End Sub

    Private Function getSqlSaldo(ByVal Parametros As Dictionary(Of String, Object))
        Dim TipoDeDocumento As eTipoDeDocumento = eTipoDeDocumento.Nota
        Dim EntradaSaida As eEntradaSaida = eEntradaSaida.Entrada
        Dim strCliente As String() = Nothing
        Dim strNota As String() = Nothing
        Dim CodigoProduto As String = String.Empty
        Dim DataInicio As String = String.Empty
        Dim DataFim As String = String.Empty
        Dim Pedido As Integer = 0

        If Parametros.ContainsKey("Empresa") Then
            ViewState("Empresa") = Parametros("Empresa")
        End If

        If Parametros.ContainsKey("Cliente") Then
            strCliente = Parametros("Cliente")
        End If

        If Parametros.ContainsKey("TipoDeDocumento") Then
            TipoDeDocumento = Parametros("TipoDeDocumento")
        End If

        If Parametros.ContainsKey("EntradaSaida") Then
            EntradaSaida = Parametros("EntradaSaida")
        End If

        If Parametros.ContainsKey("Nota") Then
            strNota = Parametros("Nota")
        End If

        If Parametros.ContainsKey("Produto") Then
            CodigoProduto = Parametros("Produto")
        End If

        If Parametros.ContainsKey("DataInicio") Then
            DataInicio = Parametros("DataInicio")
        End If

        If Parametros.ContainsKey("DataFim") Then
            DataFim = Parametros("DataFim")
        End If
        If Parametros.ContainsKey("Pedido") Then
            Pedido = Parametros("Pedido")
        End If

        Dim Sql As String = "SELECT NF.Empresa_id, " & vbCrLf & _
                            "       NF.EndEmpresa_Id, " & vbCrLf & _
                            "       NF.Cliente_Id, " & vbCrLf & _
                            "       NF.EndCliente_Id, " & vbCrLf & _
                            "       Clientes.Nome, " & vbCrLf & _
                            "       NFxI.sequencia_id, " & vbCrLf & _
                            "       NFxI.cfop_id, " & vbCrLf & _
                            "       prd.Produto_Id + '-' + prd.Descricao as Produto," & vbCrLf & _
                            "       NF.EntradaSaida_Id, " & vbCrLf & _
                            "       NF.Serie_Id, " & vbCrLf & _
                            "       NF.Nota_Id, " & vbCrLf & _
                            "       NF.Operacao, " & vbCrLf & _
                            "       NF.SubOperacao, " & vbCrLf & _
                            "       NFERealizadas.ChaveNfe, " & vbCrLf & _
                            "      'False' as chkSelecionado, " & vbCrLf & _
                            "       NFERealizadas.Data, " & vbCrLf & _
                            "       NFxI.QuantidadeFiscal as PesoFiscal, " & vbCrLf & _
                            "       --ISNULL(NFI.QuantidadeFiscal,0) as PesoFiscal, " & vbCrLf & _
                            "       ISNULL(consumido.QuantidadeFiscalConsumido ,0) as ConsumidoPeso, " & vbCrLf & _
                            " " & vbCrLf & _
                            "       ISNULL(NFI.Valor, 0) as ValorOficial, " & vbCrLf & _
                            "       ISNULL(NFI.Valor, 0) / CASE WHEN ISNULL(NFI.QuantidadeFiscal, 0)> 0 THEN ISNULL(NFI.QuantidadeFiscal, 0) ELSE 1 END as UnitarioCalculado, " & vbCrLf & _
                            "       ISNULL(consumido.ValorConsumido,0) as ConsumidoValor, " & vbCrLf & _
                            " " & vbCrLf & _
                            "       ISNULL(NFI.QuantidadeFiscal,0) - ISNULL(consumido.QuantidadeFiscalConsumido ,0) as PesoSaldo, " & vbCrLf & _
                            "       ISNULL(NFI.Valor, 0)           - ISNULL(consumido.ValorConsumido,0)             as ValorSaldo, " & vbCrLf & _
                            " " & vbCrLf & _
                            "       '' as CRTC_Empresa_Id,  " & vbCrLf & _
                            "       '' as CRTC_EndEmpresa_Id,  " & vbCrLf & _
                            "       '' as CRTC_Cliente_Id,  " & vbCrLf & _
                            "       '' as CRTC_EndCliente_Id,  " & vbCrLf & _
                            "       '' as CRTC_EntradaSaida_Id,  " & vbCrLf & _
                            "       '' as CRTC_Serie_Id,  " & vbCrLf & _
                            "       '' as CRTC_Nota_Id,  " & vbCrLf & _
                            "       '' as CRTC_TipoDeDocumento,  " & vbCrLf & _
                            "       ISNULL(consumido.ValorConsumido,0) AS FreteValorTotal" & vbCrLf & _
                            "  FROM NotasFiscais NF " & vbCrLf & _
                            " Inner JOIN NotasFiscaisxItens NFxI  " & vbCrLf & _
                            "    ON NF.Empresa_Id = NFxI.Empresa_id  " & vbCrLf & _
                            "   AND NF.EndEmpresa_Id = NFxI.EndEmpresa_id  " & vbCrLf & _
                            "   AND NF.Cliente_Id = NFxI.Cliente_id  " & vbCrLf & _
                            "   AND NF.EndCliente_Id = NFxI.EndCliente_id  " & vbCrLf & _
                            "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                            "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                            "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                            " Inner Join Produtos prd" & vbCrLf & _
                            "    on prd.Produto_Id = NFxI.Produto_Id" & vbCrLf & _
                            " INNER JOIN Operacoes Op" & vbCrLf & _
                            "    ON NF.Operacao = Op.Operacao_Id" & vbCrLf & _
                            " INNER JOIN SubOperacoes SOP " & vbCrLf & _
                            "    ON NF.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf & _
                            "   AND NF.Operacao = SOP.Operacao_Id" & vbCrLf & _
                            "  JOIN Pedidos P " & vbCrLf & _
                            "    ON NF.Pedido = P.Pedido_Id " & vbCrLf & _
                            "   AND NF.Empresa_Id = P.Empresa_Id " & vbCrLf & _
                            "   AND NF.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                            "   AND NF.Cliente_Id = P.Cliente " & vbCrLf & _
                            "   AND NF.EndCliente_Id = P.EndCliente " & vbCrLf & _
                            " INNER JOIN (SELECT Empresa_Id," & vbCrLf & _
                            "                          EndEmpresa_Id," & vbCrLf & _
                            "                          Cliente_Id," & vbCrLf & _
                            "                          EndCliente_Id," & vbCrLf & _
                            "                          EntradaSaida_Id," & vbCrLf & _
                            "                          Serie_Id," & vbCrLf & _
                            "                          Nota_Id," & vbCrLf & _
                            "                          SUM(QuantidadeFiscal) AS QuantidadeFiscal," & vbCrLf & _
                            "                          SUM(Valor) AS Valor " & vbCrLf & _
                            "	                  FROM NotasFiscaisXItens " & vbCrLf & _
                            "	                 GROUP BY Empresa_id," & vbCrLf & _
                            "                          EndEmpresa_id," & vbCrLf & _
                            "                          Cliente_id," & vbCrLf & _
                            "                          EndCliente_id," & vbCrLf & _
                            "                          EntradaSaida_id," & vbCrLf & _
                            "                          Serie_id," & vbCrLf & _
                            "                          Nota_id " & vbCrLf & _
                            "                   ) As NFI  " & vbCrLf & _
                            "    ON NFI.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
                            "   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
                            "   AND NFI.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
                            "   AND NFI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
                            "   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
                            "   AND NFI.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                            "   AND NFI.Nota_Id         = NF.Nota_Id             " & vbCrLf & _
                            " INNER JOIN Clientes " & vbCrLf & _
                            "    ON NF.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf & _
                            "   AND NF.EndCliente_Id = Clientes.Endereco_Id  " & vbCrLf & _
                            "  LEFT JOIN NFERealizadas " & vbCrLf & _
                            "    ON NF.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                            "   AND NF.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                            "   AND NF.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                            "   AND NF.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                            "   AND NF.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                            "   AND NF.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                            "   AND NF.Nota_Id         = NFERealizadas.Nota_Id " & vbCrLf & _
                            "  LEFT JOIN (SELECT nxi.Empresa_Id, " & vbCrLf & _
                            "                     nxi.EndEmpresa_Id, " & vbCrLf & _
                            "                     nxi.Cliente_Id, " & vbCrLf & _
                            "                     nxi.EndCliente_Id, " & vbCrLf & _
                            "                     nxi.EntradaSaida_Id, " & vbCrLf & _
                            "                     nxi.Serie_Id, " & vbCrLf & _
                            "                     nxi.Nota_Id, " & vbCrLf & _
                            "                     SUM(nfr.Quantidade) AS QuantidadeFiscalConsumido, " & vbCrLf & _
                            "                     SUM(nfr.Valor) AS ValorConsumido " & vbCrLf & _
                            "               FROM NotasFiscaisXItens nxi " & vbCrLf & _
                            "              INNER JOIN NotaFiscalReferencial nfr " & vbCrLf & _
                            "                 ON nxi.Empresa_Id = nfr.Empresa_Id " & vbCrLf & _
                            "                AND nxi.EndEmpresa_Id = nfr.EndEmpresa_Id " & vbCrLf & _
                            "                AND nxi.Cliente_Id = nfr.Cliente_Id " & vbCrLf & _
                            "                AND nxi.EndCliente_Id = nfr.EndCliente_Id " & vbCrLf & _
                            "                AND nxi.EntradaSaida_Id = nfr.EntradaSaida_Id " & vbCrLf & _
                            "                AND nxi.Nota_Id = nfr.Nota_Id " & vbCrLf & _
                            "                AND nxi.Serie_Id = nfr.Serie_Id " & vbCrLf & _
                            "                AND nxi.Produto_Id = nfr.Produto_Id " & vbCrLf & _
                            "                AND nxi.CFOP_Id = nfr.CFOP_Id " & vbCrLf & _
                            "                AND nxi.Sequencia_Id = nfr.Sequencia_Id " & vbCrLf & _
                            "                AND nfr.TipoReferencial_Id = '" & TipoReferencial.ToString & "'" & vbCrLf & _
                            "              GROUP BY nxi.Empresa_id, " & vbCrLf & _
                            "                       nxi.EndEmpresa_id, " & vbCrLf & _
                            "                       nxi.Cliente_id, " & vbCrLf & _
                            "                       nxi.EndCliente_id, " & vbCrLf & _
                            "                       nxi.EntradaSaida_id, " & vbCrLf & _
                            "                       nxi.Serie_id, " & vbCrLf & _
                            "                       nxi.Nota_id " & vbCrLf & _
                            "                       ) As Consumido   " & vbCrLf & _
                            "    ON Consumido.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                            "   AND Consumido.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                            "   AND Consumido.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                            "   AND Consumido.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                            "   AND Consumido.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                            "   AND Consumido.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                            "   AND Consumido.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                            "  LEFT JOIN NotaFiscalDevolucaoXNotaFiscal Dev " & vbCrLf & _
                            "    ON Dev.EmpresaDevolucao_Id	 = NF.Empresa_Id " & vbCrLf & _
                            "   AND Dev.EndEmpresaDevolucao_Id = NF.EndEmpresa_Id " & vbCrLf & _
                            "   AND Dev.ClienteDevolucao_Id	 = NF.Cliente_Id " & vbCrLf & _
                            "   AND Dev.EndClienteDevolucao_Id = NF.EndCliente_Id " & vbCrLf & _
                            "   AND Dev.Nota_Id			     = NF.Nota_Id  " & vbCrLf & _
                            "   AND Dev.Serie_Id		         = NF.Serie_Id  " & vbCrLf & _
                            "   AND Dev.EntradaSaida_Id        = NF.EntradaSaida_Id " & vbCrLf & _
                            " WHERE 1=1 " & vbCrLf & _
                            "   AND (NF.Empresa_Id = '" & ViewState("Empresa")(0) & "') " & vbCrLf & _
                            "   AND (NF.EndEmpresa_Id = " & ViewState("Empresa")(1) & ")" & vbCrLf & _
                            "   AND (NF.Situacao in (1,4,7)) " & vbCrLf
        'If TipoReferencial = eTipoReferencial.RPA Then
        '    Sql &= "   AND SOP.Classe IN ('" & eClassesOperacoes.COMPRAS.ToString & "' ,'" & eClassesOperacoes.TRANSFERENCIAS.ToString & "')" & vbCrLf
        'Else
        '    Sql &= "   AND (NF.EntradaSaida_Id = '" & IIf(EntradaSaida = eEntradaSaida.Saida, "S", "E") & "') " & vbCrLf
        'End If


        If TipoReferencial <> eTipoReferencial.RPA Then
            Sql &= "   AND (NF.EntradaSaida_Id = '" & IIf(EntradaSaida = eEntradaSaida.Saida, "S", "E") & "') " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(CodigoProduto) Then
            Sql &= "   AND (NFxi.Produto_id = '" & CodigoProduto & "') " & vbCrLf
        End If

        If strCliente IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strCliente(0)) Then
            Sql &= "   AND (NF.Cliente_Id = '" & strCliente(0) & "') " & vbCrLf
            Sql &= "   AND (NF.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
        End If

        If strNota IsNot Nothing AndAlso strNota.Length > 0 Then
            Sql &= "   AND (NF.Nota_Id IN (" & String.Join(",", strNota) & ")) " & vbCrLf
        End If

        Sql &= "   AND (NF.TipoDeDocumento = " & CInt(TipoDeDocumento) & ") " & vbCrLf
        Sql &= "   AND (ISNULL(NFI.QuantidadeFiscal,0) - ISNULL(consumido.QuantidadeFiscalConsumido ,0)) > 0 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(DataInicio) And Not String.IsNullOrWhiteSpace(DataFim) Then
            Sql &= "   AND (NF.Movimento between '" & DataInicio.ToSqlDate() & "' and '" & DataFim.ToSqlDate() & "') " & vbCrLf
        End If

        If TipoReferencial = eTipoReferencial.EXP Then
            Sql &= "   AND OP.Classe ='" & eClassesOperacoes.COMPRAS.ToString & "'" & vbCrLf & _
                   "   AND YEAR(NF.Movimento) >= YEAR(GETDATE())-2 " & vbCrLf & _
                   "   AND SOP.Memorando = 1" & vbCrLf
        End If

        If Pedido > 0 Then
            Sql &= "   AND NF.Pedido=" & Pedido & vbCrLf
        End If

        If TipoReferencial = eTipoReferencial.EXP Then
            Sql &= "  ORDER BY NF.Nota_id " & vbCrLf
        Else
            Sql &= " ORDER BY NFERealizadas.Nota_id " & vbCrLf
        End If

        Return Sql

    End Function

    Private Function calculaSaldoTotal() As Decimal
        Dim i As Integer = 0
        Dim quantidadeFiscalTotal As Decimal = 0
        For Each objNFSel As [Lib].Negocio.NotaFiscal In ListaDeNotasFiscais
            quantidadeFiscalTotal = quantidadeFiscalTotal + objNFSel.Itens(0).QuantidadeFiscal
        Next
        Return quantidadeFiscalTotal
    End Function

    Private Function calculaValorTotalAtualizado() As Decimal
        Dim Valor As Decimal = 0
        Dim ValorTotal As Decimal = 0

        For Each row As GridViewRow In grdNFeSaldo.Rows
            If row.RowType = DataControlRowType.DataRow Then
                If TipoReferencial = eTipoReferencial.RPA Then
                    Valor = CDec(CType(grdNFeSaldo.Rows(row.RowIndex).FindControl("txtFreteValorTotal"), TextBox).Text)
                Else
                    Valor = CDec(CType(grdNFeSaldo.Rows(row.RowIndex).FindControl("txtValorAtualizado"), TextBox).Text)
                End If
                ValorTotal = ValorTotal + Valor
            End If
        Next
        Return ValorTotal
    End Function

    Private Function calculaPesoAjustavel() As Decimal
        Dim Quantidade As Decimal
        Dim QuantidadeTotal As Decimal

        If TipoReferencial = eTipoReferencial.RPA Then
            Dim lbl As Label
            For Each row As GridViewRow In grdNFeSaldo.Rows
                If row.RowType = DataControlRowType.DataRow Then
                    lbl = grdNFeSaldo.Rows(row.RowIndex).FindControl("lblPesoFiscal")
                    Quantidade = CInt(lbl.Text)
                    QuantidadeTotal = QuantidadeTotal + Quantidade
                End If
            Next
        Else
            Dim Txt As TextBox
            For Each row As GridViewRow In grdNFeSaldo.Rows
                If row.RowType = DataControlRowType.DataRow Then
                    Txt = grdNFeSaldo.Rows(row.RowIndex).FindControl("txtVlrPeso")
                    Quantidade = Txt.Text
                    QuantidadeTotal = QuantidadeTotal + Quantidade
                End If
            Next
        End If
        Return QuantidadeTotal
    End Function

    Private Sub calculaTotaisSelecionados()
        Dim txtPeso As TextBox
        Dim txtValor As TextBox
        Dim lblPeso As Label
        Dim chkSelecionado As CheckBox

        Dim pesoTotal As Decimal = 0
        Dim valorTotal As Decimal = 0

        For Each row As GridViewRow In grdNFeSaldo.Rows
            If row.RowType = DataControlRowType.DataRow Then
                chkSelecionado = CType(grdNFeSaldo.Rows(row.RowIndex).FindControl("chkSelecionado"), CheckBox)
                If chkSelecionado.Checked Then
                    If TipoReferencial = eTipoReferencial.RPA Then
                        lblPeso = grdNFeSaldo.Rows(row.RowIndex).FindControl("lblPesoFiscal")
                        txtValor = grdNFeSaldo.Rows(row.RowIndex).FindControl("txtFreteValorTotal")
                        valorTotal = valorTotal + CDec(txtValor.Text)
                        pesoTotal = pesoTotal + CInt(lblPeso.Text)
                    Else
                        txtPeso = grdNFeSaldo.Rows(row.RowIndex).FindControl("txtVlrPeso")
                        txtValor = grdNFeSaldo.Rows(row.RowIndex).FindControl("txtValorAtualizado")
                        pesoTotal = pesoTotal + CInt(txtPeso.Text)
                        valorTotal = valorTotal + CDec(txtValor.Text)
                    End If
                End If
            End If
        Next
        txtPesoTotalSelecionado.Text = CInt(pesoTotal)
        txtValorTotalSelecionado.Text = CDec(valorTotal)
    End Sub

#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divNotaFiscalReferencial")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            Dim lstSelecionados As List(Of Int32) = grdNFeSaldo.GetSelectedItems("chkSelecionado")

            If TipoReferencial = eTipoReferencial.RPA Then
                For i As Integer = 0 To grdNFeSaldo.Rows.Count - 1
                    If CType(grdNFeSaldo.Rows(i).FindControl("chkSelecionado"), CheckBox).Checked AndAlso CDec(CType(grdNFeSaldo.Rows(i).FindControl("txtFreteValorTotal"), TextBox).Text) = 0 Then
                        MsgBox(Me.Page, "Existem Notas Selecionadas sem Valor de Frete!")
                        Exit Sub
                    End If
                Next
            End If

            Dim lstNotasFiscaisReferenciais As New ListNotaFiscalReferencial()
            Dim objNotaFiscalReferencial As NotaFiscalReferencial

            For Each row As GridViewRow In grdNFeSaldo.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                If chk.Checked Then
                    objNotaFiscalReferencial = New NotaFiscalReferencial()
                    objNotaFiscalReferencial.IUD = "I"

                    objNotaFiscalReferencial.Empresa_Id = ViewState("Empresa")(0)
                    objNotaFiscalReferencial.EndEmpresa_Id = ViewState("Empresa")(1)
                    objNotaFiscalReferencial.Cliente_Id = row.Cells(4).Text
                    objNotaFiscalReferencial.EndCliente_Id = row.Cells(5).Text
                    objNotaFiscalReferencial.EntradaSaida_Id = IIf(row.Cells(1).Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    objNotaFiscalReferencial.Serie_Id = row.Cells(3).Text
                    objNotaFiscalReferencial.Nota_Id = row.Cells(2).Text
                    objNotaFiscalReferencial.Produto_Id = row.Cells(7).Text.Split("-")(0)
                    objNotaFiscalReferencial.Sequencia_Id = row.Cells(8).Text
                    objNotaFiscalReferencial.CFOP_Id = row.Cells(9).Text
                    objNotaFiscalReferencial.Quantidade = CType(row.FindControl("lblPesoFiscal"), Label).Text  'quantidade da nota
                    objNotaFiscalReferencial.Valor = CType(row.FindControl("txtFreteValorTotal"), TextBox).Text  'valor do frete
                    objNotaFiscalReferencial.TipoReferencial_Id = TipoReferencial
                    lstNotasFiscaisReferenciais.Add(objNotaFiscalReferencial)
                End If
            Next

            If TipoReferencial = eTipoReferencial.EXP AndAlso lstNotasFiscaisReferenciais.QuantidadeTotal > QuantidadeFiscalNF Then
                MsgBox(Me.Page, "A Quantidade Total das Nf's Referenciadas não pode ser maior do que a NF de Exportação!")
            ElseIf TipoReferencial = eTipoReferencial.RPA AndAlso lstNotasFiscaisReferenciais.ValorTotal <> ValorNFOrigem Then
                MsgBox(Me.Page, String.Format("O valor total de Frete das notas selecionadas {0} está diferente do valor do RPA {1} !", lstNotasFiscaisReferenciais.ValorTotal.ToString("N2"), ValorNFOrigem.ToString("N2")))
            Else
                Session(Session("ssTipoRetorno")) = lstNotasFiscaisReferenciais
                If Session("ssTipoRetorno") IsNot Nothing Then
                    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                        Dim ucName = MainUserControl.ClientID.Split("_")
                        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                        CType(uc, IBaseUserControl).Carregar(lstNotasFiscaisReferenciais)
                    Else
                        CType(Me.Page, IBasePage).Carregar(lstNotasFiscaisReferenciais)
                    End If
                    Popup.CloseDialog(Me.Page, "divNotaFiscalReferencial")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtVlrPeso_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim txtVlrPeso As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txtVlrPeso.NamingContainer, GridViewRow)

            Dim txtValorAtualizado As TextBox = grdNFeSaldo.Rows(row.RowIndex).FindControl("txtValorAtualizado")
            Dim quantidadeFiscalAnterior As Integer = CInt(ListaDeNotasFiscais(row.RowIndex).Itens(0).QuantidadeFiscal)

            If Not String.IsNullOrWhiteSpace(txtVlrPeso.Text) AndAlso CInt(txtVlrPeso.Text) > 0 Then
                ListaDeNotasFiscais(row.RowIndex).Itens(0).QuantidadeFiscal = CInt(txtVlrPeso.Text)
                Dim valorUnitario As Decimal = CDec(ListaDeNotasFiscais(row.RowIndex).Itens(0).ValorTotal) / CInt(quantidadeFiscalAnterior)
                ListaDeNotasFiscais(row.RowIndex).Itens(0).ValorTotal = Math.Round(CInt(txtVlrPeso.Text) * CDec(valorUnitario), 2, MidpointRounding.AwayFromZero)
                txtValorAtualizado.Text = ListaDeNotasFiscais(row.RowIndex).Itens(0).ValorTotal
            End If

            txtPesoTotal.Text = calculaPesoAjustavel()
            txtTotalValorAtualizado.Text = calculaValorTotalAtualizado()
            calculaTotaisSelecionados()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub grdNFeSaldo_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grdNFeSaldo.RowDataBound
        If e.Row.RowType = DataControlRowType.Header OrElse e.Row.RowType = DataControlRowType.DataRow Then
            If TipoReferencial = eTipoReferencial.RPA Then
                e.Row.Cells(15).Visible = False
                e.Row.Cells(16).Visible = False
                e.Row.Cells(17).Visible = False
                e.Row.Cells(18).Visible = False
                e.Row.Cells(19).Visible = False

            Else
                e.Row.Cells(20).Visible = False
            End If
        End If
    End Sub

    Protected Sub txtFreteValorTotal_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim txtFreteValorTotal As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txtFreteValorTotal.NamingContainer, GridViewRow)

            Dim Chk As CheckBox = grdNFeSaldo.Rows(row.RowIndex).FindControl("chkSelecionado")

            If Chk.Checked Then
                If Not String.IsNullOrWhiteSpace(txtFreteValorTotal.Text) AndAlso CInt(txtFreteValorTotal.Text) > 0 Then
                    ListaDeNotasFiscais(row.RowIndex).Itens(0).ValorTotal = CDec(txtFreteValorTotal.Text)
                End If
                txtFreteValorTotal.Enabled = True
            Else
                ListaDeNotasFiscais(row.RowIndex).Itens(0).ValorTotalMoeda = 0.0
                ListaDeNotasFiscais(row.RowIndex).Itens(0).ValorTotal = 0.0
                txtFreteValorTotal.Enabled = False
            End If

            calculaTotaisSelecionados()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

#End Region

End Class